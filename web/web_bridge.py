# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Cầu nối cho bản chạy trong trình duyệt (Pyodide).

Nhận (method, url, body) từ JavaScript, chuyển cho `app.api_core.xu_ly` (dùng chung với FastAPI)
và trả về (status, headers_json, body)."""
from __future__ import annotations

import json
from urllib.parse import parse_qs, urlsplit

from app import api_core, ban_quyen

ban_quyen.kiem_tra()
MAX_BYTES = 300 * 1024 * 1024


def handle(method, url, body_js):
    """Điểm vào duy nhất cho JavaScript. Luôn trả (status, headers_json, bytes)."""
    try:
        body = bytes(body_js.to_py()) if body_js is not None else b""
        if len(body) > MAX_BYTES:
            r = api_core._json({"detail": "File quá lớn cho bản chạy trong trình duyệt."}, 413)
        else:
            u = urlsplit(url)
            q = {k: v[0] for k, v in parse_qs(u.query, keep_blank_values=True).items()}
            r = api_core.xu_ly(method, u.path, q, body)
        return r.status, json.dumps(r.headers), r.body
    except Exception as e:
        return 500, json.dumps({"Content-Type": "application/json"}), \
            json.dumps({"detail": f"Lỗi xử lý: {type(e).__name__}: {e}"}, ensure_ascii=False).encode("utf-8")
