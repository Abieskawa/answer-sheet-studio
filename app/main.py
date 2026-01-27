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
        "footer_reopen_hint": "若看到「127.0.0.1 拒絕連線」，請重新執行 start_mac.command（macOS）或 start_windows.vbs（Windows）啟動。",
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
        "download_hint_support": "年級欄位支援 7–12（含 10/11/12），班級支援 1–9（0 保留），座號支援 00–99，選擇題支援 ABC / ABCD / ABCDE，最多 100 題。",
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
        "upload_hint_output": "完成後會輸出 results.csv、annotated.pdf，以及答案分析報表/圖表。",
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
        "result_download_analysis_pdf": "下載分析結果 PDF",
        "result_download_showwrong": "下載 showwrong.xlsx（只顯示錯題）",
        "result_print_pdf": "列印／存成 PDF",
        "result_print_pdf_hint": "用瀏覽器列印（⌘/Ctrl+P）：macOS 選「PDF → 另存為 PDF」；Windows 選「Microsoft Print to PDF」。",
        "result_back_to_downloads": "查看試題分析原始數據",
        "result_view_item_analysis_data": "試題分析數據",
        "result_plots_title": "圖表",
        "result_plot_score_hist": "成績分佈",
        "result_plot_score_hist_guides": "輔助線：實線＝平均（紅）/中位數（藍）；虛線＝百分位數（P12/P25/P75/P88）",
        "result_plot_item_metrics": "題目分析",
        "result_plot_item_metrics_guides": "輔助線：難度 0.25/0.50/0.75；鑑別度 0.00/0.20/0.40",
        "result_integrated_title": "整合圖表（可用滑鼠懸停查看題號，並與錯題表對照）",
        "result_integrated_hint": "滑鼠移到表格/圖表可同步醒目提示題號；點一下可鎖定，再點一次取消。",
        "result_integrated_error": "讀取整合資料失敗",
        "result_interaction_click_to_lock": "點擊可鎖定/取消醒目提示",
        "result_controls_filter_student": "篩選學生",
        "result_controls_filter_student_ph": "搜尋座號或 ID",
        "result_controls_sort": "排序",
        "result_controls_sort_seat": "座號",
        "result_controls_sort_score_desc": "分數（高→低）",
        "result_controls_sort_blank_asc": "空白（少→多）",
        "result_controls_sort_id": "ID",
        "result_controls_hide_correct": "隱藏正確",
        "result_controls_charts_follow_filter": "圖表跟著篩選",
        "result_controls_jump_question": "跳到題號",
        "result_controls_jump_question_ph": "題號",
        "result_controls_go": "前往",
        "result_controls_clear": "清除醒目/鎖定",
        "result_controls_not_found": "找不到",
        "result_focus_no_selection": "把滑鼠移到表格/圖表可查看",
        "result_focus_answer": "作答",
        "result_focus_score": "分數",
        "result_focus_open_page": "開啟該生標註頁",
        "result_focus_locked": "鎖定",
        "result_focus_keyboard_hint": "快捷鍵：←/→切換題號、↑/↓切換學生、Esc 清除鎖定、PgUp/PgDn 切換圖",
        "result_focus_status_correct": "正確",
        "result_focus_status_wrong": "錯誤",
        "result_focus_status_blank": "空白",
        "result_loading": "載入中…",
        "result_chart_tab_difficulty": "難度",
        "result_chart_tab_discrimination": "鑑別度",
        "result_chart_tab_scores": "成績分佈",
        "result_chart_tab_blank_rate": "空白率",
        "result_chart_prev": "上一張",
        "result_chart_next": "下一張",
        "col_question": "題號",
        "col_correct": "正確答案",
        "col_student": "學生",
        "col_difficulty": "難度（正答率）",
        "col_discrimination": "鑑別度",
        "col_class": "班級",
        "col_wrong": "錯誤人數",
        "col_blank": "空白人數",
        "col_wrong_short": "錯",
        "col_blank_short": "空",
        "result_analysis_files_title": "試題分析檔案",
        "result_analysis_files_hint": "下載完整分析報表（CSV/XLSX/圖片/日誌）。",
        "result_item_table_title": "試題分析數據（預覽）",
        "result_item_table_hint": "此處僅顯示前 200 列；完整資料請下載 analysis_item.csv。",
        "result_item_table_truncated": "（共 {total_rows} 列，顯示前 {shown_rows} 列）",
        "result_download_scores_by_class": "analysis_scores_by_class.xlsx（按班級）",
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
        "analysis_error_builtin_failed": "內建分析失敗：",
        "analysis_message_done": "分析完成，可下載報表與圖表。",
        "analysis_message_done_fallback": "分析完成。",
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
        "footer_reopen_hint": "If you see “127.0.0.1 refused to connect”, rerun start_mac.command (macOS) or start_windows.vbs (Windows) to start the local server.",
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
        "download_hint_support": "Grade supports 7–12 (including 10/11/12). Class supports 1–9 (0 reserved). Seat No supports 00–99. Choices support ABC / ABCD / ABCDE. Up to 100 questions.",
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
        "upload_hint_output": "Outputs results.csv, annotated.pdf, and analysis reports/plots.",
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
        "result_download_analysis_pdf": "Download analysis report PDF",
        "result_download_showwrong": "Download showwrong.xlsx (wrong answers only)",
        "result_print_pdf": "Print / Save as PDF",
        "result_print_pdf_hint": "Print via your browser (Cmd/Ctrl+P): macOS choose “PDF → Save as PDF”; Windows choose “Microsoft Print to PDF”.",
        "result_back_to_downloads": "View item analysis raw data",
        "result_view_item_analysis_data": "Item analysis data",
        "result_plots_title": "Plots",
        "result_plot_score_hist": "Score distribution",
        "result_plot_score_hist_guides": "Guide lines: solid = mean (red) / median (blue); dashed = percentiles (P12/P25/P75/P88)",
        "result_plot_item_metrics": "Item analysis",
        "result_plot_item_metrics_guides": "Guide lines: difficulty 0.25/0.50/0.75; discrimination 0.00/0.20/0.40",
        "result_integrated_title": "Integrated report (hover a point to highlight the corresponding row)",
        "result_integrated_hint": "Hover the table/charts to highlight a question; click to lock/unlock.",
        "result_integrated_error": "Failed to load integrated data",
        "result_interaction_click_to_lock": "Click to lock/unlock highlight",
        "result_controls_filter_student": "Filter student",
        "result_controls_filter_student_ph": "Search seat or ID",
        "result_controls_sort": "Sort",
        "result_controls_sort_seat": "Seat",
        "result_controls_sort_score_desc": "Score (high→low)",
        "result_controls_sort_blank_asc": "Blanks (low→high)",
        "result_controls_sort_id": "ID",
        "result_controls_hide_correct": "Hide correct",
        "result_controls_charts_follow_filter": "Charts follow filter",
        "result_controls_jump_question": "Jump to question",
        "result_controls_jump_question_ph": "Qno",
        "result_controls_go": "Go",
        "result_controls_clear": "Clear highlight/locks",
        "result_controls_not_found": "Not found",
        "result_focus_no_selection": "Hover the table/charts to inspect",
        "result_focus_answer": "Answer",
        "result_focus_score": "Score",
        "result_focus_open_page": "Open annotated page",
        "result_focus_locked": "locked",
        "result_focus_keyboard_hint": "Shortcuts: ←/→ question, ↑/↓ student, Esc clear locks, PgUp/PgDn switch chart",
        "result_focus_status_correct": "Correct",
        "result_focus_status_wrong": "Wrong",
        "result_focus_status_blank": "Blank",
        "result_loading": "Loading…",
        "result_chart_tab_difficulty": "Difficulty",
        "result_chart_tab_discrimination": "Discrimination",
        "result_chart_tab_scores": "Score distribution",
        "result_chart_tab_blank_rate": "Blank rate",
        "result_chart_prev": "Prev",
        "result_chart_next": "Next",
        "col_question": "Question",
        "col_correct": "Correct",
        "col_student": "Student",
        "col_difficulty": "Difficulty (accuracy)",
        "col_discrimination": "Discrimination",
        "col_class": "Class",
        "col_wrong": "Wrong",
        "col_blank": "Blank",
        "col_wrong_short": "Wrong",
        "col_blank_short": "Blank",
        "result_analysis_files_title": "Item analysis files",
        "result_analysis_files_hint": "Download the full analysis reports (CSV/XLSX/images/logs).",
        "result_item_table_title": "Item analysis data (preview)",
        "result_item_table_hint": "Only the first 200 rows are shown here; download analysis_item.csv for the full data.",
        "result_item_table_truncated": "({total_rows} rows, showing first {shown_rows})",
        "result_download_scores_by_class": "analysis_scores_by_class.xlsx (by class)",
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
        "analysis_error_builtin_failed": "Built-in analysis failed:",
        "analysis_message_done": "Analysis complete. Download reports and plots below.",
        "analysis_message_done_fallback": "Analysis complete.",
        "analysis_note_discrimination_rule_27": "Discrimination: if students > 30, uses top/bottom 27% (correct_high − correct_low).",
        "analysis_note_discrimination_rule_50": "Discrimination: if students ≤ 30, uses top/bottom 50% (correct_high − correct_low).",
    },
}

def _normalize_i18n() -> None:
    base = I18N.get(DEFAULT_LANG) or {}
    all_keys: set[str] = set(base.keys())
    for lang, table in I18N.items():
        all_keys |= set((table or {}).keys())
        if not isinstance(table, dict):
            I18N[lang] = {}
    for lang, table in I18N.items():
        for k in all_keys:
            if k not in table:
                table[k] = base.get(k, k)


_normalize_i18n()

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
    analysis_report_pdf = job_dir / "analysis_report.pdf"
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
            "analysis_report_url": (f"/outputs/{job_id}/analysis_report.pdf" if analysis_report_pdf.exists() else None),
            "analysis_error": (str(meta.get("analysis_error") or "") or None),
            "analysis_message": (str(meta.get("analysis_message") or "") or None),
            "analysis_discrimination_note_key": discr_note_key,
            "analysis_score_hist_inline_url": (f"/outputs_inline/{job_id}/analysis_score_hist.png" if score_hist.exists() else None),
            "analysis_item_plot_inline_url": (f"/outputs_inline/{job_id}/analysis_item_plot.png" if item_plot.exists() else None),
            "analysis_files": _analysis_file_links(job_id),
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
    template_csv = job_dir / "analysis_template.csv"
    roster_csv = job_dir / "roster.csv"
    item_csv = job_dir / "analysis_item.csv"
    scores_csv = job_dir / "analysis_scores.csv"
    analysis_report_pdf = job_dir / "analysis_report.pdf"
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
            "analysis_template_csv_url": (f"/outputs/{job_id}/analysis_template.csv" if template_csv.exists() else None),
            "roster_csv_url": (f"/outputs/{job_id}/roster.csv" if roster_csv.exists() else None),
            "analysis_item_csv_url": (f"/outputs/{job_id}/analysis_item.csv" if item_csv.exists() else None),
            "analysis_scores_csv_url": (f"/outputs/{job_id}/analysis_scores.csv" if scores_csv.exists() else None),
            "analysis_report_url": (f"/outputs/{job_id}/analysis_report.pdf" if analysis_report_pdf.exists() else None),
            "analysis_item_table": item_table,
            "analysis_files": _analysis_file_links(job_id),
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
        "analysis_template.csv": "analysis_template.csv",
        "analysis_scores.csv": "analysis_scores.csv",
        "analysis_scores_by_class.xlsx": "analysis_scores_by_class.xlsx",
        "analysis_item.csv": "analysis_item.csv",
        "analysis_summary.csv": "analysis_summary.csv",
        "analysis_score_hist.png": "analysis_score_hist.png",
        "analysis_item_plot.png": "analysis_item_plot.png",
    }

    def add_if_exists(name: str) -> None:
        path = job_dir / name
        try:
            if not path.exists():
                return
        except Exception:
            return
        files.append({"url": f"/outputs/{job_id}/{name}", "label": label_by_name.get(name, name)})

    # Keep a stable, friendly order for known files first.
    for name in label_by_name.keys():
        add_if_exists(name)

    # Then include any extra analysis outputs (future-proofing).
    try:
        for path in sorted(job_dir.iterdir()):
            name = path.name
            if name in label_by_name:
                continue
            if not name.startswith("analysis_"):
                continue
            if path.is_dir():
                continue
            if not any(name.lower().endswith(ext) for ext in (".csv", ".xlsx", ".png", ".log", ".txt")):
                continue
            add_if_exists(name)
    except Exception:
        pass

    return files


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

    # Also emit a lightweight JSON for the interactive HTML report.
    try:
        classes: list[dict] = []
        if grouped_students:
            for key in sorted(grouped_students.keys()):
                students = grouped_students[key][:]
                students.sort(key=student_sort_key)
                rows_payload: list[dict] = []
                for i, qno in enumerate(q_numbers):
                    correct = corrects[i] if i < len(corrects) else ""
                    if not correct:
                        cells = ["" for _ in students]
                    else:
                        cells = []
                        for student in students:
                            answers = answers_by_student.get(student, [])
                            ans = answers[i] if i < len(answers) else ""
                            cells.append("" if ans == correct else (blank_label if ans == "" else ans))
                    rows_payload.append({"number": qno, "correct": correct, "cells": cells})
                classes.append({"key": key, "label": key, "students": students, "rows": rows_payload})
        else:
            rows_payload: list[dict] = []
            for i, qno in enumerate(q_numbers):
                correct = corrects[i] if i < len(corrects) else ""
                if not correct:
                    cells = ["" for _ in student_cols]
                else:
                    cells = []
                    for student in student_cols:
                        answers = answers_by_student.get(student, [])
                        ans = answers[i] if i < len(answers) else ""
                        cells.append("" if ans == correct else (blank_label if ans == "" else ans))
                rows_payload.append({"number": qno, "correct": correct, "cells": cells})
            classes.append({"key": "all", "label": "all", "students": student_cols, "rows": rows_payload})

        out_json = Path(out_xlsx_path).with_name("analysis_showwrong.json")
        out_json.write_text(json.dumps({"classes": classes}, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def _write_analysis_report_pdf(job_dir: Path, lang: str) -> None:
    import math

    template_csv = Path(job_dir) / "analysis_template.csv"
    if not template_csv.exists():
        return

    t = I18N.get(lang, I18N[DEFAULT_LANG])

    try:
        with open(template_csv, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            fieldnames = [str(x or "").strip() for x in (reader.fieldnames or [])]
            required = ("number", "correct", "points")
            if not all(k in fieldnames for k in required):
                return
            student_cols = [c for c in fieldnames if c not in required and c]
            template_rows = list(reader)
    except Exception:
        return

    if not student_cols or not template_rows:
        return

    score_hist_path = Path(job_dir) / "analysis_score_hist.png"
    item_plot_path = Path(job_dir) / "analysis_item_plot.png"

    q_numbers: list[int] = []
    corrects: list[str] = []
    points: list[float] = []
    answers_by_student: dict[str, list[str]] = {s: [] for s in student_cols}

    for row in template_rows:
        raw_no = str(row.get("number") or "").strip()
        try:
            qno = int(float(raw_no))  # tolerate Excel-like "1.0"
        except Exception:
            continue
        q_numbers.append(qno)
        corrects.append(_normalize_answer_cell(row.get("correct") or ""))
        raw_points = str(row.get("points") or "").strip()
        try:
            p = float(raw_points) if raw_points != "" else 0.0
        except Exception:
            p = 0.0
        points.append(float(p))
        for s in student_cols:
            answers_by_student[s].append(_normalize_answer_cell(row.get(s) or ""))

    if not q_numbers:
        return

    # Roster (optional): used for class grouping & ordering.
    roster_path = Path(job_dir) / "roster.csv"
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

    def seat_number(meta: dict[str, str]) -> Optional[int]:
        raw = (meta.get("seat_no") or "").strip()
        if not raw:
            return None
        try:
            return int(float(raw))
        except Exception:
            return None

    def student_sort_key(student_id: str) -> tuple[int, int, str]:
        meta = roster_by_person.get(student_id, {})
        seat = seat_number(meta)
        if seat is None:
            return (1, 10**9, str(student_id))
        return (0, seat, str(student_id))

    grouped_students: dict[str, list[str]] = {}
    for s in student_cols:
        key = class_key(roster_by_person.get(s, {})) if roster_by_person else "all"
        grouped_students.setdefault(key, []).append(s)
    if not grouped_students:
        grouped_students = {"all": student_cols}

    groups = [(k, sorted(grouped_students[k], key=student_sort_key)) for k in sorted(grouped_students.keys())]

    def compute_scores(students: list[str]) -> dict[str, float]:
        scores: dict[str, float] = {}
        for s in students:
            ans = answers_by_student.get(s, [])
            score = 0.0
            for i in range(len(q_numbers)):
                c = corrects[i] if i < len(corrects) else ""
                if not c:
                    continue
                a = ans[i] if i < len(ans) else ""
                if a and a == c:
                    score += float(points[i] if i < len(points) else 0.0)
            scores[s] = float(score)
        return scores

    def compute_item_metrics(
        students: list[str],
        scores_by_student: dict[str, float],
    ) -> tuple[list[Optional[float]], list[Optional[float]], list[Optional[float]], list[int], list[int]]:
        n = len(students)
        diffs: list[Optional[float]] = [None for _ in q_numbers]
        discs: list[Optional[float]] = [None for _ in q_numbers]
        blank_rates: list[Optional[float]] = [None for _ in q_numbers]
        wrong_counts: list[int] = [0 for _ in q_numbers]
        blank_counts: list[int] = [0 for _ in q_numbers]
        if n <= 0:
            return diffs, discs, blank_rates, wrong_counts, blank_counts

        for i in range(len(q_numbers)):
            c = corrects[i] if i < len(corrects) else ""
            if not c:
                continue
            ok = 0
            wrong = 0
            blank = 0
            for s in students:
                ans = answers_by_student.get(s, [])
                a = ans[i] if i < len(ans) else ""
                if not a:
                    blank += 1
                elif a == c:
                    ok += 1
                else:
                    wrong += 1
            diffs[i] = ok / n
            blank_rates[i] = blank / n
            wrong_counts[i] = int(wrong)
            blank_counts[i] = int(blank)

        group_n = int(math.floor(n * 0.27)) if n > 30 else int(math.floor(n / 2.0))
        if group_n > 0:
            ordered = sorted(students, key=lambda sid: (scores_by_student.get(sid, 0.0), str(sid)))
            low = set(ordered[:group_n])
            high = set(ordered[-group_n:])
            for i in range(len(q_numbers)):
                c = corrects[i] if i < len(corrects) else ""
                if not c:
                    continue
                low_ok = 0
                high_ok = 0
                for s in low:
                    ans = answers_by_student.get(s, [])
                    a = ans[i] if i < len(ans) else ""
                    if a and a == c:
                        low_ok += 1
                for s in high:
                    ans = answers_by_student.get(s, [])
                    a = ans[i] if i < len(ans) else ""
                    if a and a == c:
                        high_ok += 1
                discs[i] = (high_ok / group_n) - (low_ok / group_n)

        return diffs, discs, blank_rates, wrong_counts, blank_counts

    def fmt2(v: Optional[float]) -> str:
        return "" if v is None else f"{float(v):.2f}"

    def safe_token(value: str) -> str:
        token = (value or "").strip().replace("\x00", "")
        token = re.sub(r"[^A-Za-z0-9-]+", "_", token)
        token = re.sub(r"_+", "_", token).strip("_.-")
        return token or "all"

    def make_metric_plot_png(
        numbers: list[int],
        values: list[Optional[float]],
        out_path: Path,
        title_text: str,
        *,
        y_min: Optional[float],
        y_max: Optional[float],
        ref_lines: list[float],
        bgr_color: tuple[int, int, int],
    ) -> bool:
        try:
            import cv2  # type: ignore
            import numpy as np  # type: ignore
        except (ImportError, OSError):
            return False

        width, height = 1200, 300
        left, right, top, bottom = 70, 30, 40, 45
        plot_w = width - left - right
        plot_h = height - top - bottom

        img = np.full((height, width, 3), 255, dtype=np.uint8)
        axis = (0, 0, 0)
        grid = (229, 231, 235)  # #e5e7eb

        xs = [float(x) for x in numbers]
        ys = [None if v is None else float(v) for v in values]
        if not xs:
            return False

        x_min = float(min(xs))
        x_max = float(max(xs))
        if x_max <= x_min:
            x_max = x_min + 1.0

        y_vals = [v for v in ys if v is not None and math.isfinite(v)]
        if y_min is None or y_max is None:
            if y_vals:
                y_min = float(min(y_vals))
                y_max = float(max(y_vals))
                pad = max(0.1, (y_max - y_min) * 0.1)
                y_min -= pad
                y_max += pad
            else:
                y_min, y_max = 0.0, 1.0
        if y_max <= y_min:
            y_max = y_min + 1.0

        def x_of(x: float) -> int:
            return int(left + (x - x_min) / (x_max - x_min) * plot_w)

        def y_of(y: float) -> int:
            return int(top + (1.0 - (y - y_min) / (y_max - y_min)) * plot_h)

        # Axes
        cv2.line(img, (left, top), (left, top + plot_h), axis, 2)
        cv2.line(img, (left, top + plot_h), (left + plot_w, top + plot_h), axis, 2)

        # Reference lines
        for y in ref_lines:
            if y_min <= y <= y_max:
                py = y_of(float(y))
                cv2.line(img, (left, py), (left + plot_w, py), grid, 1)

        # Line + points
        last: Optional[tuple[int, int]] = None
        for x, y in zip(xs, ys):
            if y is None or not math.isfinite(float(y)):
                last = None
                continue
            pt = (x_of(float(x)), y_of(float(y)))
            if last is not None:
                cv2.line(img, last, pt, (209, 213, 219), 2)  # #d1d5db
            cv2.circle(img, pt, 4, bgr_color, thickness=-1)
            last = pt

        # Title + labels
        cv2.putText(img, str(title_text), (left, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.9, axis, 2, cv2.LINE_AA)
        cv2.putText(img, str(t.get("col_question", "Question")), (left + plot_w // 2 - 24, height - 16), cv2.FONT_HERSHEY_SIMPLEX, 0.7, axis, 2, cv2.LINE_AA)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        cv2.imwrite(str(out_path), img)
        return True

    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether
    except Exception:
        return

    styles = getSampleStyleSheet()
    base_font = "Helvetica"
    header_font = "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        if str(lang).lower().startswith("zh"):
            base_font = "STSong-Light"
            header_font = "STSong-Light"
    except Exception:
        pass

    try:
        styles["Title"].fontName = base_font
        styles["Heading2"].fontName = base_font
        styles["BodyText"].fontName = base_font
    except Exception:
        pass

    doc = SimpleDocTemplate(
        str(Path(job_dir) / "analysis_report.pdf"),
        pagesize=landscape(A4),
        leftMargin=24,
        rightMargin=24,
        topMargin=22,
        bottomMargin=22,
    )
    story: list = []

    title = t.get("result_download_analysis_pdf") or "Analysis report"
    story.append(Paragraph(str(title), styles["Title"]))
    story.append(Spacer(1, 10))

    def add_cover_plot(path: Path, heading: str, aspect_wh: tuple[float, float]) -> bool:
        try:
            if not path.exists():
                return False
        except Exception:
            return False
        try:
            flow = [Paragraph(str(heading), styles["Heading2"]), Spacer(1, 6)]

            w0, h0 = float(aspect_wh[0]), float(aspect_wh[1])
            if w0 <= 0 or h0 <= 0:
                w0, h0 = 1.0, 1.0

            max_w = float(doc.width)
            # Keep generous vertical room for the heading and spacing so the image never spills
            # into the next page (which would separate the heading from the plot).
            max_h = float(doc.height) - 80.0

            w = max_w
            h = w * (h0 / w0)
            if h > max_h:
                h = max_h
                w = h * (w0 / h0)

            flow.append(Image(str(path), width=w, height=h))
            story.append(KeepTogether(flow))
            story.append(PageBreak())
            return True
        except Exception:
            return False

    # Cover plots (overall): score histogram + item plot.
    add_cover_plot(
        score_hist_path,
        str(t.get("result_plot_score_hist", "Score distribution")),
        (800.0, 450.0),
    )
    add_cover_plot(
        item_plot_path,
        str(t.get("result_plot_item_metrics", "Item analysis")),
        (800.0, 700.0),
    )

    metric_specs = [
        (
            "difficulty",
            str(t.get("result_chart_tab_difficulty", "Difficulty")),
            0.0,
            1.0,
            [0.25, 0.5, 0.75],
            (74, 222, 128),  # BGR-ish green
        ),
        (
            "discrimination",
            str(t.get("result_chart_tab_discrimination", "Discrimination")),
            None,
            None,
            [0.0, 0.2, 0.4],
            (235, 99, 37),  # BGR-ish blue-ish
        ),
        (
            "blank_rate",
            str(t.get("result_chart_tab_blank_rate", "Blank rate")),
            0.0,
            1.0,
            [0.1, 0.2, 0.3],
            (128, 128, 128),
        ),
    ]

    def build_compact_rows(
        per_row: int,
        diffs: list[Optional[float]],
        discs: list[Optional[float]],
        blank_rates: list[Optional[float]],
        wrong_counts: list[int],
        blank_counts: list[int],
    ) -> list[list[str]]:
        cols_per_q = [
            t.get("col_question", "Q"),
            t.get("result_chart_tab_difficulty", t.get("col_difficulty", "Difficulty")),
            t.get("result_chart_tab_discrimination", t.get("col_discrimination", "Discrimination")),
            t.get("result_chart_tab_blank_rate", "Blank rate"),
            t.get("col_wrong_short", t.get("col_wrong", "Wrong")),
            t.get("col_blank_short", t.get("col_blank", "Blank")),
        ]
        header = [*cols_per_q] * per_row
        rows: list[list[str]] = [header]

        def one_q(i: int) -> list[str]:
            return [
                str(q_numbers[i]),
                fmt2(diffs[i]),
                fmt2(discs[i]),
                fmt2(blank_rates[i]),
                str(int(wrong_counts[i])),
                str(int(blank_counts[i])),
            ]

        i = 0
        while i < len(q_numbers):
            row: list[str] = []
            for j in range(per_row):
                idx = i + j
                if idx < len(q_numbers):
                    row.extend(one_q(idx))
                else:
                    row.extend([""] * len(cols_per_q))
            rows.append(row)
            i += per_row
        return rows

    for class_idx, (label, students) in enumerate(groups):
        if class_idx > 0:
            story.append(PageBreak())
            story.append(Paragraph(f"{t.get('col_class', 'Class')}: {label}", styles["Title"]))
            story.append(PageBreak())

        scores_by_student = compute_scores(students)
        diffs, discs, blank_rates, wrong_counts, blank_counts = compute_item_metrics(students, scores_by_student)

        per_row = 2 if len(q_numbers) <= 40 else 3
        table_rows = build_compact_rows(per_row, diffs, discs, blank_rates, wrong_counts, blank_counts)

        for metric_idx, (metric_key, metric_label, y_min, y_max, ref_lines, bgr) in enumerate(metric_specs):
            if metric_idx > 0:
                story.append(PageBreak())

            story.append(Paragraph(f"{t.get('col_class', 'Class')}: {label} ({len(students)}) · {metric_label}", styles["Heading2"]))
            story.append(Spacer(1, 6))

            col_count = len(table_rows[0]) if table_rows else 0
            col_widths = None
            if col_count > 0:
                # 6 columns per question chunk: Q | Diff | Disc | BlankRate | Wrong | Blank
                chunk = [28, 48, 54, 54, 44, 44] if per_row <= 2 else [24, 40, 44, 44, 34, 34]
                col_widths = chunk * per_row

            tbl = Table(table_rows, repeatRows=1, colWidths=col_widths)
            tbl.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                        ("FONTNAME", (0, 0), (-1, 0), header_font),
                        ("FONTSIZE", (0, 0), (-1, 0), 8),
                        ("FONTNAME", (0, 1), (-1, -1), base_font),
                        ("FONTSIZE", (0, 1), (-1, -1), 7),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 2),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                        ("TOPPADDING", (0, 0), (-1, -1), 1),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                    ]
                )
            )
            story.append(tbl)
            story.append(Spacer(1, 6))

            metric_values = {
                "difficulty": diffs,
                "discrimination": discs,
                "blank_rate": blank_rates,
            }.get(metric_key, diffs)

            token = safe_token(label)
            chart_path = Path(job_dir) / f"analysis_report_chart_{token}_{metric_key}.png"
            if make_metric_plot_png(q_numbers, metric_values, chart_path, metric_label, y_min=y_min, y_max=y_max, ref_lines=ref_lines, bgr_color=bgr):
                chart_w = float(doc.width)
                chart_h = chart_w * (300.0 / 1200.0)
                story.append(Image(str(chart_path), width=chart_w, height=chart_h))

    doc.build(story)


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
            try:
                from engine.analysis import run_analysis_template

                run_analysis_template(template_path, job_dir, lang=lang)
                analysis_message = t.get("analysis_message_done", "Analysis complete.")
                analysis_error = None
            except Exception as exc:
                analysis_error = f"{t.get('analysis_error_builtin_failed', 'Built-in analysis failed:')} {exc}".strip()
            else:
                try:
                    _write_analysis_report_pdf(job_dir, lang=lang)
                except Exception:
                    pass

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
    elif filename == "analysis_showwrong.json":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_showwrong.json"
    elif filename == "analysis_report.pdf":
        download_name = f"{upload_base_ascii}_{job_tag}_analysis_report.pdf"

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
