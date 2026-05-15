from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Depends, Request, Response
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from functools import lru_cache
from pathlib import Path
from pydantic import BaseModel
from urllib import error as urllib_error
from urllib import request as urllib_request
import shutil
import tempfile
import os
import traceback
import json
from typing import Optional

import engine
from api.frontend_bundle import EMBEDDED_INDEX_HTML

app = FastAPI(title="Xiaomi Daily Report Engine API")

def first_env(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


SUPABASE_URL = first_env(
    "NEXT_PUBLIC_SUPABASE_URL",
    "SUPABASE_URL",
).rstrip("/")
SUPABASE_PUBLISHABLE_KEY = first_env(
    "NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY",
    "NEXT_PUBLIC_SUPABASE_ANON_KEY",
    "SUPABASE_PUBLISHABLE_KEY",
    "SUPABASE_ANON_KEY",
)
AUTH_COOKIE_NAME = "zopper-access-token"
ALLOWED_EMAIL_DOMAIN = "@zopper.com"
COOKIE_SECURE = bool(os.getenv("VERCEL"))

# Mount static files
STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
INDEX_FILE = STATIC_DIR / "index.html"
STYLES_FILE = STATIC_DIR / "styles.css"
APP_JS_FILE = STATIC_DIR / "app.js"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class LoginPayload(BaseModel):
    email: str
    password: str


def normalise_email(email: str) -> str:
    return email.strip().lower()


def is_allowed_email(email: str) -> bool:
    normalised = normalise_email(email)
    local_part, separator, domain = normalised.partition("@")
    return bool(local_part and separator and f"@{domain}" == ALLOWED_EMAIL_DOMAIN)


def ensure_supabase_config() -> None:
    if not SUPABASE_URL or not SUPABASE_PUBLISHABLE_KEY:
        raise HTTPException(
            500,
            "Supabase authentication is not configured. Set SUPABASE_URL and SUPABASE_ANON_KEY or their NEXT_PUBLIC_* equivalents.",
        )


def supabase_request(
    path: str,
    *,
    method: str = "GET",
    payload: Optional[dict[str, object]] = None,
    access_token: Optional[str] = None,
) -> dict[str, object]:
    ensure_supabase_config()

    headers = {"apikey": SUPABASE_PUBLISHABLE_KEY}
    body = None
    if payload is not None:
        headers["Content-Type"] = "application/json"
        body = json.dumps(payload).encode("utf-8")
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"

    request = urllib_request.Request(
        f"{SUPABASE_URL}{path}",
        data=body,
        headers=headers,
        method=method,
    )

    try:
        with urllib_request.urlopen(request) as response:
            raw = response.read().decode("utf-8") or "{}"
            return json.loads(raw)
    except urllib_error.HTTPError as exc:
        raw = exc.read().decode("utf-8") if exc.fp else ""
        message = "Authentication failed."
        if raw:
            try:
                parsed = json.loads(raw)
                message = (
                    parsed.get("msg")
                    or parsed.get("error_description")
                    or parsed.get("message")
                    or parsed.get("error")
                    or message
                )
            except json.JSONDecodeError:
                message = raw
        raise HTTPException(401 if exc.code in {400, 401, 403} else 502, message)
    except urllib_error.URLError:
        raise HTTPException(502, "Could not reach Supabase authentication service.")


def fetch_authenticated_user(access_token: str) -> dict[str, object]:
    return supabase_request("/auth/v1/user", access_token=access_token)


def set_auth_cookie(response: Response, access_token: str, expires_in: int) -> None:
    response.set_cookie(
        AUTH_COOKIE_NAME,
        access_token,
        max_age=expires_in,
        httponly=True,
        secure=COOKIE_SECURE,
        samesite="lax",
        path="/",
    )


def clear_auth_cookie(response: Response) -> None:
    response.delete_cookie(AUTH_COOKIE_NAME, path="/")


async def get_authenticated_user(request: Request) -> dict[str, object]:
    access_token = request.cookies.get(AUTH_COOKIE_NAME)
    if not access_token:
        raise HTTPException(401, "Please sign in.")

    user = fetch_authenticated_user(access_token)
    email = str(user.get("email", ""))
    if not is_allowed_email(email):
        raise HTTPException(403, "Only @zopper.com accounts are allowed.")
    return user


@lru_cache(maxsize=1)
def load_frontend_html() -> str:
    if not (INDEX_FILE.exists() and STYLES_FILE.exists() and APP_JS_FILE.exists()):
        return EMBEDDED_INDEX_HTML

    html = INDEX_FILE.read_text(encoding="utf-8")
    html = html.replace(
        '<link rel="stylesheet" href="/static/styles.css">',
        f"<style>\n{STYLES_FILE.read_text(encoding='utf-8')}\n</style>",
    )
    html = html.replace(
        '<script src="/static/app.js"></script>',
        f"<script>\n{APP_JS_FILE.read_text(encoding='utf-8')}\n</script>",
    )
    return html

@app.get("/", response_class=HTMLResponse)
async def read_index():
    return load_frontend_html()

@app.post("/api/auth/login")
async def login(payload: LoginPayload, response: Response):
    email = normalise_email(payload.email)
    password = payload.password.strip()

    if not is_allowed_email(email):
        raise HTTPException(400, "Use your @zopper.com email address.")
    if not password:
        raise HTTPException(400, "Password is required.")

    session = supabase_request(
        "/auth/v1/token?grant_type=password",
        method="POST",
        payload={"email": email, "password": password},
    )
    access_token = str(session.get("access_token", ""))
    if not access_token:
        raise HTTPException(401, "Login failed.")

    user = session.get("user") or fetch_authenticated_user(access_token)
    if not is_allowed_email(str(user.get("email", ""))):
        raise HTTPException(403, "Only @zopper.com accounts are allowed.")

    set_auth_cookie(response, access_token, int(session.get("expires_in", 3600)))
    return {"user": {"email": str(user.get("email", email))}}


@app.post("/api/auth/logout")
async def logout(response: Response):
    clear_auth_cookie(response)
    return {"ok": True}


@app.get("/api/auth/session")
async def auth_session(user: dict[str, object] = Depends(get_authenticated_user)):
    return {"user": {"email": str(user.get("email", ""))}}


@app.get("/api/status")
async def get_status(user: dict[str, object] = Depends(get_authenticated_user)):
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
    user: dict[str, object] = Depends(get_authenticated_user),
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
async def download_get(file_key: str, user: dict[str, object] = Depends(get_authenticated_user)):
    return await handle_download_logic(file_key)

@app.post("/api/download")
async def download_post(
    report_type: str = Form(...),
    download_key: str = Form(...),
    service_file: Optional[UploadFile] = File(None),
    axio_file: Optional[UploadFile] = File(None),
    retail_file: Optional[UploadFile] = File(None),
    user: dict[str, object] = Depends(get_authenticated_user),
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
