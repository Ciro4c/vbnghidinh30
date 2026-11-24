# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Web app chuẩn hóa văn bản theo NĐ 30/2020/NĐ-CP và slide chuẩn.

Chạy:  python -m uvicorn app.main:app --reload
"""
import os
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

from . import api_core, ban_quyen

STATIC_DIR = Path(__file__).parent / "static"
try:                                    # chạy trên máy cá nhân nên cho phép file lớn; chỉnh bằng biến môi trường
    MAX_MB = max(1, int(os.environ.get("VANBAN_MAX_MB", "500")))
except ValueError:
    MAX_MB = 500
MAX_BYTES = MAX_MB * 1024 * 1024

ban_quyen.kiem_tra()

app = FastAPI(title="Chuẩn hóa văn bản NĐ 30/2020 và slide chuẩn")


@app.middleware("http")
async def them_tac_gia(request: Request, call_next):
    resp = await call_next(request)
    resp.headers["X-Tac-Gia"] = ban_quyen.CHU_KY
    return resp

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


CHAN_TRANG = (
    '<footer id="chu-ky" style="text-align:center;padding:18px 12px 28px;font:13px system-ui,sans-serif;'
    'opacity:.7">{ban_quyen} · <span>{chu_ky}</span></footer>'
)


@app.get("/")
def trang_chu():
    ban_quyen.kiem_tra()
    html = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    chan = CHAN_TRANG.format(ban_quyen=ban_quyen.BAN_QUYEN, chu_ky=ban_quyen.CHU_KY)
    html = html.replace("</body>", chan + "\n</body>") if "</body>" in html else html + chan
    return HTMLResponse(html)


@app.api_route("/api/{duong:path}", methods=["GET", "POST"])
async def api(duong: str, request: Request):
    """Mọi API đi qua api_core.xu_ly (dùng chung với bản chạy trong trình duyệt)."""
    body = b""
    if request.method == "POST":
        body = await request.body()
        if len(body) > MAX_BYTES:
            raise HTTPException(413, f"File quá lớn (tối đa {MAX_MB} MB). Có thể nâng giới hạn bằng biến "
                                     "môi trường VANBAN_MAX_MB.")
    q = {k: v for k, v in request.query_params.items()}
    r = await run_in_threadpool(api_core.xu_ly, request.method, "/api/" + duong, q, body)
    h = {k: v for k, v in r.headers.items() if k.lower() not in ("content-type",)}
    return Response(content=r.body, status_code=r.status, media_type=r.headers.get("Content-Type"), headers=h)
