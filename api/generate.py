from __future__ import annotations

import sys
import traceback
from pathlib import Path
from http.server import BaseHTTPRequestHandler

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    from app import XiaomiReportHandler as handler
except Exception as e:
    error_traceback = traceback.format_exc()
    class handler(BaseHTTPRequestHandler):
        def do_POST(self):
            self.send_response(500)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(f"Import Error in generate.py:\n{error_traceback}".encode("utf-8"))
