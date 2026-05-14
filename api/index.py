from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from functools import lru_cache
from pathlib import Path
import shutil
import tempfile
import os
import traceback
from typing import Optional

import engine

app = FastAPI(title="Xiaomi Daily Report Engine API")

# Mount static files
STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
INDEX_FILE = STATIC_DIR / "index.html"
STYLES_FILE = STATIC_DIR / "styles.css"
APP_JS_FILE = STATIC_DIR / "app.js"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@lru_cache(maxsize=1)
def load_frontend_html() -> str:
    if not INDEX_FILE.exists():
        return "<h1>Xiaomi Daily Report Engine</h1><p>Frontend not found.</p>"

    html = INDEX_FILE.read_text(encoding="utf-8")
    if STYLES_FILE.exists():
        html = html.replace(
            '<link rel="stylesheet" href="/static/styles.css">',
            f"<style>\n{STYLES_FILE.read_text(encoding='utf-8')}\n</style>",
        )
    if APP_JS_FILE.exists():
        html = html.replace(
            '<script src="/static/app.js"></script>',
            f"<script>\n{APP_JS_FILE.read_text(encoding='utf-8')}\n</script>",
        )
    return html

@app.get("/", response_class=HTMLResponse)
async def read_index():
    return load_frontend_html()

@app.get("/api/status")
async def get_status():
    try:
        return {
            "app_name": engine.APP_NAME,
            "defaults": {
                "service": engine.file_status(engine.DEFAULT_SERVICE_FILE),
                "axio": engine.file_status(engine.DEFAULT_AXIO_FILE),
                "retail": engine.file_status(engine.DEFAULT_RETAIL_FILE),
                "service_master": engine.file_status(engine.DEFAULT_SERVICE_MASTER_FILE),
                "channel_master": engine.file_status(engine.DEFAULT_CHANNEL_MASTER_FILE),
            },
            "outputs": {
                "final_report": engine.file_status(engine.FINAL_REPORT_FILE),
                "zonal_report": engine.file_status(engine.ZONAL_REPORT_FILE),
                "channel_report": engine.file_status(engine.CHANNEL_REPORT_FILE),
            }
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "traceback": traceback.format_exc()})

@app.post("/api/generate")
async def generate(
    report_type: str = Form(...),
    service_file: Optional[UploadFile] = File(None),
    axio_file: Optional[UploadFile] = File(None),
    retail_file: Optional[UploadFile] = File(None),
):
    try:
        result = await run_generation_logic(report_type, service_file, axio_file, retail_file)
        return {
            "active_report": report_type,
            "reports": {report_type: result},
            "summary": result["summary"],
            "preview": result["preview"],
            "columns": result["columns"],
            "downloads": result["downloads"]
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "traceback": traceback.format_exc()})

@app.get("/download/{file_key}.xlsx")
async def download_get(file_key: str):
    return await handle_download_logic(file_key)

@app.post("/api/download")
async def download_post(
    report_type: str = Form(...),
    download_key: str = Form(...),
    service_file: Optional[UploadFile] = File(None),
    axio_file: Optional[UploadFile] = File(None),
    retail_file: Optional[UploadFile] = File(None),
):
    try:
        # Optimization: If the file already exists in /tmp (from a recent generation), just serve it.
        targets = {
            "final_report": engine.FINAL_REPORT_FILE,
            "zonal_report": engine.ZONAL_REPORT_FILE,
            "channel_report": engine.CHANNEL_REPORT_FILE,
        }
        
        target_path = targets.get(download_key)
        if target_path and target_path.exists():
            return await handle_download_logic(download_key)
            
        await run_generation_logic(report_type, service_file, axio_file, retail_file)
        return await handle_download_logic(download_key)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "traceback": traceback.format_exc()})

async def run_generation_logic(
    report_type: str,
    service_file: Optional[UploadFile],
    axio_file: Optional[UploadFile],
    retail_file: Optional[UploadFile],
):
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        if report_type == "service":
            s_path = await save_upload(service_file, tmp_path) or engine.DEFAULT_SERVICE_FILE
            m_path = engine.DEFAULT_SERVICE_MASTER_FILE
            if not s_path.exists(): raise HTTPException(400, "Service CSV file missing")
            if not m_path.exists(): raise HTTPException(400, f"Master file not found in repository: {m_path.name}")
            result = engine.generate_service_report(s_path, m_path)
            result["downloads"] = {"final_report": "final_report"}
            return result
        elif report_type == "channel":
            a_path = await save_upload(axio_file, tmp_path) or engine.DEFAULT_AXIO_FILE
            r_path = await save_upload(retail_file, tmp_path) or engine.DEFAULT_RETAIL_FILE
            m_path = engine.DEFAULT_CHANNEL_MASTER_FILE
            if not a_path.exists(): raise HTTPException(400, "Axio CSV file missing")
            if not r_path.exists(): raise HTTPException(400, "Retail CSV file missing")
            if not m_path.exists(): raise HTTPException(400, f"Master file not found in repository: {m_path.name}")
            result = engine.generate_channel_payload(a_path, r_path, m_path)
            result["downloads"] = {"channel_report": "channel_report"}
            return result
        raise HTTPException(400, "Invalid report type")

async def handle_download_logic(file_key: str):
    targets = {
        "final_report": (engine.FINAL_REPORT_FILE, "final_report.xlsx"),
        "channel_report": (engine.CHANNEL_REPORT_FILE, "final_channel_report.xlsx"),
    }
    if file_key not in targets:
        raise HTTPException(404, "File key not found")
    
    path, filename = targets[file_key]
    if not path.exists():
        raise HTTPException(404, "File not found on server. Please generate it first.")
        
    return FileResponse(path, filename=filename, media_type=engine.EXCEL_CONTENT_TYPE)

async def save_upload(upload: Optional[UploadFile], target_dir: Path) -> Optional[Path]:
    if not upload or not upload.filename: return None
    path = target_dir / upload.filename
    with path.open("wb") as buffer:
        shutil.copyfileobj(upload.file, buffer)
    return path
