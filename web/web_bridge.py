# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Cầu nối cho bản chạy trong trình duyệt (Pyodide).

Thay cho FastAPI: nhận (method, url, body) từ JavaScript, gọi đúng hàm xử lý trong `app/`
và trả về (status, headers_json, body). Logic nghiệp vụ dùng lại nguyên từ app/.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from pydantic import ValidationError

from app import ban_quyen, lo as lo_mod
from app.importers.docx_reader import doc_docx_ex, doc_txt
from app.importers.phan_tich import phan_tich
from app.importers.pptx_reader import doc_pptx
from app.profile import DEFAULT_PROFILE
from app.renderers.cong_van import cong_van_to_bytes
from app.renderers.slide_hoc_thuat import ThongTin, dung_slide, ten_profile
from app.schemas import TEN_LOAI, CongVan

ban_quyen.kiem_tra()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_BYTES = 200 * 1024 * 1024


class Loi(Exception):
    def __init__(self, status: int, detail):
        self.status, self.detail = status, detail


def _json(obj, status=200, headers=None):
    h = {"Content-Type": "application/json", "X-Tac-Gia": ban_quyen.CHU_KY}
    h.update(headers or {})
    return status, json.dumps(h), json.dumps(obj, ensure_ascii=False).encode("utf-8")


def _file(data: bytes, mime: str, ten: str, extra=None):
    h = {"Content-Type": mime, "Content-Disposition": f'attachment; filename="{ten}"',
         "X-Tac-Gia": ban_quyen.CHU_KY}
    h.update(extra or {})
    return 200, json.dumps(h), data


def _dau_header(warns):
    return json.dumps([w[:300] for w in warns[:15]], ensure_ascii=True)


def _kiem_tra_body(body: bytes) -> bytes:
    if not body:
        raise Loi(400, "Chưa nhận được nội dung file.")
    if len(body) > MAX_BYTES:
        raise Loi(413, "File quá lớn cho bản chạy trong trình duyệt.")
    return body


def _import(q, body):
    data = _kiem_tra_body(body)
    ten = q.get("ten", "")
    ext = Path(ten).suffix.lower()
    try:
        hinh: dict = {}
        if ext == ".docx":
            lines, w0, hinh = doc_docx_ex(data)
        elif ext == ".txt":
            lines, w0 = doc_txt(data)
        elif ext == ".doc":
            raise Loi(400, "File .doc cũ chưa đọc được. Hãy mở bằng Word, chọn Lưu thành (Save As) "
                           "định dạng .docx rồi tải lại.")
        else:
            raise Loi(400, "Chỉ nhận file .docx hoặc .txt (PowerPoint: dùng tab Slide).")
    except Loi:
        raise
    except Exception:
        raise Loi(400, "Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu).")
    if not lines:
        raise Loi(400, "File không có nội dung văn bản.")
    fields, w1 = phan_tich(lines)
    return _json({"fields": fields, "canh_bao": w0 + w1, "hinh": hinh})


def _cong_van(q, body):
    try:
        data = CongVan(**json.loads(body.decode("utf-8")))
    except ValidationError as e:
        raise Loi(422, [{"loc": ["body", *map(str, x["loc"])], "msg": x["msg"]} for x in e.errors()])
    except ValueError:
        raise Loi(422, "Dữ liệu gửi lên không hợp lệ.")
    try:
        content = cong_van_to_bytes(data, q.get("profile", DEFAULT_PROFILE))
    except (FileNotFoundError, ValueError, KeyError) as e:
        raise Loi(400, str(e))
    return _file(content, DOCX_MIME, "van-ban.docx")


def _slide(q, body):
    data = _kiem_tra_body(body)
    ext = Path(q.get("ten", "")).suffix.lower()
    if ext == ".ppt":
        raise Loi(400, "File .ppt cũ chưa đọc được. Hãy mở bằng PowerPoint, chọn Lưu thành định dạng .pptx rồi tải lại.")
    if ext != ".pptx":
        raise Loi(400, "Chỉ nhận file .pptx.")
    try:
        slides, w0 = doc_pptx(data)
    except Exception:
        raise Loi(400, "Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu).")
    if not slides:
        raise Loi(400, "File không có slide nào có nội dung.")
    g = lambda k: q.get(k, "").strip()
    out, w1, n = dung_slide(slides, ThongTin(g("ten_bai"), g("hoc_phan"), g("giang_vien"), g("don_vi"), g("nam")),
                            profile=ten_profile(q.get("kieu", "chuan")))
    return _file(out, PPTX_MIME, "slide-chuan.pptx",
                 {"X-Canh-Bao": _dau_header(w0 + w1), "X-So-Slide": str(n)})


def handle(method, url, body_js):
    """Điểm vào duy nhất cho JavaScript. Luôn trả (status, headers_json, bytes)."""
    try:
        body = bytes(body_js.to_py()) if body_js is not None else b""
        u = urlsplit(url)
        q = {k: v[0] for k, v in parse_qs(u.query, keep_blank_values=True).items()}
        p = u.path
        if method == "GET" and p == "/api/loai-van-ban":
            return _json({k: (v or "Công văn") for k, v in TEN_LOAI.items()})
        if method == "GET" and p == "/api/tac-gia":
            return _json({"chu_ky": ban_quyen.CHU_KY, "ban_quyen": ban_quyen.BAN_QUYEN})
        if method == "POST" and p == "/api/import":
            return _import(q, body)
        if method == "POST" and p == "/api/cong-van/docx":
            return _cong_van(q, body)
        if method == "POST" and p == "/api/slide/chuan-hoa":
            return _slide(q, body)
        if method == "POST" and p == "/api/lo/tao":
            return _json({"lo": lo_mod.tao_lo()})
        m = re.fullmatch(r"/api/lo/([^/]+)/them", p)
        if method == "POST" and m:
            _kiem_tra_body(body)
            tt = {k: q.get(k, "") for k in ("hoc_phan", "giang_vien", "don_vi", "nam")}
            tt["kieu"] = q.get("kieu", "chuan")
            try:
                return _json(lo_mod.xu_ly_tep(m.group(1), q.get("ten", ""), body, tt))
            except (ValueError, FileNotFoundError) as e:
                raise Loi(400, str(e))
        m = re.fullmatch(r"/api/lo/([^/]+)/zip", p)
        if method == "GET" and m:
            try:
                data = lo_mod.tao_zip(m.group(1)).read_bytes()
            except (ValueError, FileNotFoundError) as e:
                raise Loi(404, str(e))
            return _file(data, "application/zip", "ket-qua-chuan-hoa.zip")
        raise Loi(404, "Không tìm thấy đường dẫn.")
    except Loi as e:
        return _json({"detail": e.detail}, e.status)
    except Exception as e:  # lỗi bất ngờ: báo về giao diện thay vì làm treo
        return _json({"detail": f"Lỗi xử lý: {type(e).__name__}: {e}"}, 500)
