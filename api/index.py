from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
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
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

@app.get("/", response_class=HTMLResponse)
async def read_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return index_file.read_text(encoding="utf-8")
    return "<h1>Xiaomi Daily Report Engine</h1><p>Frontend not found.</p>"

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
    master_file: Optional[UploadFile] = File(None),
):
    try:
        result = await run_generation_logic(report_type, service_file, axio_file, retail_file, master_file)
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
    master_file: Optional[UploadFile] = File(None),
):
    try:
        await run_generation_logic(report_type, service_file, axio_file, retail_file, master_file)
        return await handle_download_logic(download_key)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e), "traceback": traceback.format_exc()})

async def run_generation_logic(
    report_type: str,
    service_file: Optional[UploadFile],
    axio_file: Optional[UploadFile],
    retail_file: Optional[UploadFile],
    master_file: Optional[UploadFile],
):
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        if report_type == "service":
            s_path = await save_upload(service_file, tmp_path) or engine.DEFAULT_SERVICE_FILE
            m_path = await save_upload(master_file, tmp_path) or engine.DEFAULT_SERVICE_MASTER_FILE
            if not s_path.exists(): raise HTTPException(400, "Service file missing")
            if not m_path.exists(): raise HTTPException(400, "Master file missing")
            result = engine.generate_service_report(s_path, m_path)
            result["downloads"] = {"final_report": "final_report"}
            return result
        elif report_type == "channel":
            a_path = await save_upload(axio_file, tmp_path) or engine.DEFAULT_AXIO_FILE
            r_path = await save_upload(retail_file, tmp_path) or engine.DEFAULT_RETAIL_FILE
            m_path = await save_upload(master_file, tmp_path) or engine.DEFAULT_CHANNEL_MASTER_FILE
            if not a_path.exists(): raise HTTPException(400, "Axio file missing")
            if not r_path.exists(): raise HTTPException(400, "Retail file missing")
            if not m_path.exists(): raise HTTPException(400, "Master file missing")
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
