# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Lớp xử lý API dùng chung cho FastAPI (máy cá nhân) và cầu nối Pyodide (trình duyệt).

`xu_ly(method, path, query, body)` -> `Resp`. Không phụ thuộc FastAPI nên chạy được trong Pyodide.
Tham số phức tạp (tuy_chinh, tieu_de_moi...) gửi dạng JSON trong query hoặc, khi có tệp đi kèm,
trong khung VBN1 (xem khung.py)."""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from pydantic import ValidationError

from . import ban_quyen, khung, ky_tu_la, lo as lo_mod, slide_lint
from .importers.docx_reader import doc_docx_ex, doc_txt
from .importers.phan_tich import phan_tich
from .importers.pptx_reader import doc_pptx
from .profile import DEFAULT_PROFILE, ap_dung_tuy_chinh, list_profiles, load_profile, mo_ta_profile
from .renderers.cong_van import cong_van_to_bytes_ex
from .renderers.slide_hoc_thuat import ThongTin, dung_slide, ten_profile
from .schemas import TEN_LOAI, CongVan

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
PDF_MIME = "application/pdf"


class Loi(Exception):
    def __init__(self, status: int, detail):
        super().__init__(str(detail))
        self.status, self.detail = status, detail


@dataclass
class Resp:
    status: int = 200
    headers: dict = field(default_factory=dict)
    body: bytes = b""


def _json(obj, status=200, headers=None) -> Resp:
    h = {"Content-Type": "application/json"}
    h.update(headers or {})
    return Resp(status, h, json.dumps(obj, ensure_ascii=False).encode("utf-8"))


def _file(data: bytes, mime: str, ten: str, extra=None) -> Resp:
    h = {"Content-Type": mime, "Content-Disposition": f'attachment; filename="{ten}"'}
    h.update(extra or {})
    return Resp(200, h, data)


def dau_header(ds) -> str:
    return json.dumps([str(w)[:300] for w in list(ds)[:15]], ensure_ascii=True)


def _json_q(q: dict, khoa: str, mac_dinh):
    v = q.get(khoa)
    if v in (None, ""):
        return mac_dinh
    try:
        return json.loads(v)
    except ValueError:
        raise Loi(400, f"Tham số “{khoa}” không phải JSON hợp lệ.")


def _thong_tin_tu_query(q: dict) -> ThongTin:
    g = lambda k: str(q.get(k, "")).strip()
    return ThongTin(g("ten_bai"), g("hoc_phan"), g("giang_vien"), g("don_vi"), g("nam"))


def _tach_khung(body: bytes, q: dict) -> tuple[dict, dict[str, bytes]]:
    """Thân là khung VBN1 thì lấy tham số và tệp từ đó; không thì thân là tệp thô, tham số lấy ở query."""
    if khung.la_khung(body):
        try:
            return khung.tach(body)
        except Exception:
            raise Loi(400, "Gói dữ liệu gửi lên bị lỗi.")
    return dict(q), {"tep": body}


# --------------------------------------------------------------- Word
def doc_tho_docx(data: bytes) -> str:
    """Toàn bộ chữ thô của .docx (đoạn + ô bảng), CHƯA làm sạch, để dò ký tự lạ."""
    from docx import Document
    doc = Document(io.BytesIO(data))
    ds = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                ds.append(c.text)
    return "\n".join(ds)


def _import(q: dict, body: bytes) -> Resp:
    if not body:
        raise Loi(400, "Chưa nhận được nội dung file.")
    ext = Path(q.get("ten", "")).suffix.lower()
    tho = ""
    try:
        hinh: dict = {}
        if ext == ".docx":
            lines, w0, hinh = doc_docx_ex(body)
            tho = doc_tho_docx(body)
        elif ext == ".txt":
            lines, w0 = doc_txt(body)
            for enc in ("utf-8-sig", "utf-16", "cp1258"):
                try:
                    tho = body.decode(enc)
                    break
                except UnicodeError:
                    continue
        elif ext == ".pdf":
            from . import pdf_tools
            r = pdf_tools.mo_pdf(body)
            tho_trang = pdf_tools.trich_van_ban(r)
            if sum(len(t.strip()) for t in tho_trang[:5]) <= 30:
                raise Loi(400, "PDF không có lớp chữ (dạng ảnh/scan) nên không trích được nội dung. "
                               "Hãy dùng tab PDF để gắn dấu trang, watermark, chữ ký/mộc.")
            tho = "\n".join(tho_trang)
            from .importers.docx_reader import Line, lam_sach
            lines = [Line(lam_sach(x)) for x in tho.splitlines() if lam_sach(x)]
            w0 = ["Nội dung lấy từ PDF nên cấu trúc (đoạn, bảng) có thể chưa chính xác; hãy rà soát."]
        elif ext == ".doc":
            raise Loi(400, "File .doc cũ chưa đọc được. Hãy mở bằng Word, chọn Lưu thành (Save As) "
                           "định dạng .docx rồi tải lại.")
        else:
            raise Loi(400, "Chỉ nhận file .docx, .pdf hoặc .txt (PowerPoint: dùng tab Slide).")
    except Loi:
        raise
    except Exception:
        raise Loi(400, "Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu).")
    if not lines:
        raise Loi(400, "File không có nội dung văn bản.")
    la = ky_tu_la.quet(tho) if tho else []
    da_don: dict = {}
    if q.get("lam_sach") in ("1", "true"):
        tc = _json_q(q, "tuy_chon_lam_sach", None)
        moi = []
        for ln in lines:
            if ln.kind == "text":
                t, tk = ky_tu_la.lam_sach_dong(ln.text, tc)
                for k, v in tk.items():
                    da_don[k] = da_don.get(k, 0) + v
                if not t.strip():
                    continue
                ln.text = t
            moi.append(ln)
        lines = moi
    fields, w1 = phan_tich(lines)
    canh = list(w0) + list(w1) + ky_tu_la.tom_tat(la, da_don)
    if ext == ".docx":                          # đối chiếu đầy đủ: mọi từ của file gốc phải còn trong kết quả
        from .importers import doi_chieu
        canh = doi_chieu.canh_bao(body, fields) + canh
    return _json({"fields": fields, "canh_bao": canh, "hinh": hinh,
                  "ky_tu_la": [d.to_dict() for d in la[:200]], "da_don": da_don})


def _cong_van(q: dict, body: bytes) -> Resp:
    try:
        data = CongVan(**json.loads(body.decode("utf-8")))
    except ValidationError as e:
        raise Loi(422, [{"loc": ["body", *map(str, x["loc"])], "msg": x["msg"]} for x in e.errors()])
    except ValueError:
        raise Loi(422, "Dữ liệu gửi lên không hợp lệ.")
    try:
        content, canh = cong_van_to_bytes_ex(data, q.get("profile", DEFAULT_PROFILE))
    except (FileNotFoundError, ValueError, KeyError) as e:
        raise Loi(400, str(e))
    return _file(content, DOCX_MIME, "van-ban.docx", {"X-Canh-Bao": dau_header(canh)})


def _ra_soat(q: dict, body: bytes) -> Resp:
    from .ra_soat import ra_soat
    try:
        d = json.loads(body.decode("utf-8"))
    except ValueError:
        raise Loi(422, "Dữ liệu gửi lên không hợp lệ.")
    pid = d.get("profile") or DEFAULT_PROFILE
    if pid not in list_profiles():
        raise Loi(404, "Không có thể thức này.")
    return _json({"ket_qua": ra_soat(d.get("fields") or {}, pid, d.get("thanh_phan"), d.get("tuy_chinh"))})


def _ky_tu_la(q: dict, body: bytes) -> Resp:
    try:
        d = json.loads(body.decode("utf-8"))
    except ValueError:
        raise Loi(422, "Dữ liệu gửi lên không hợp lệ.")
    text = str(d.get("text", ""))
    che_do = d.get("che_do", "form")
    ds = ky_tu_la.quet(text, che_do=che_do)
    out = {"phat_hien": [p.to_dict() for p in ds], "tom_tat": ky_tu_la.tom_tat(ds), "cac_loai": ky_tu_la.dem_theo_loai(ds)}
    if d.get("lam_sach"):
        sach, tk = ky_tu_la.lam_sach(text, d.get("tuy_chon"), che_do)
        out["van_ban_sach"], out["da_don"] = sach, tk
    return _json(out)


# --------------------------------------------------------------- PDF
def _pdf_phan_tich(q: dict, body: bytes) -> Resp:
    from . import pdf_tools
    if not body:
        raise Loi(400, "Chưa nhận được file PDF.")
    try:
        return _json(pdf_tools.phan_tich(body, q.get("profile", DEFAULT_PROFILE), _json_q(q, "thanh_phan", None)))
    except Loi:
        raise
    except Exception:
        raise Loi(400, "Không đọc được PDF (có thể bị hỏng hoặc đặt mật khẩu).")


def _pdf_xu_ly(q: dict, body: bytes) -> Resp:
    from . import pdf_tools
    if not khung.la_khung(body):
        raise Loi(400, "Cần gửi gói dữ liệu (JSON + tệp).")
    meta, tep = _tach_khung(body, q)
    pdf = tep.pop("pdf", None)
    if not pdf:
        raise Loi(400, "Thiếu file PDF trong gói dữ liệu.")
    try:
        out, canh, info = pdf_tools.xu_ly_pdf(pdf, meta.get("tuy_chon") or {}, tep, meta.get("profile", DEFAULT_PROFILE))
    except ValueError as e:
        raise Loi(400, str(e))
    except Loi:
        raise
    except Exception as e:
        raise Loi(400, f"Không xử lý được PDF: {type(e).__name__}")
    return _file(out, PDF_MIME, "van-ban-da-xu-ly.pdf",
                 {"X-Canh-Bao": dau_header(canh), "X-Thong-Tin": json.dumps(info, ensure_ascii=True, default=str)[:4000]})


# --------------------------------------------------------------- slide
def _doc_pptx(data: bytes, ten: str):
    ext = Path(ten).suffix.lower()
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
    return slides, w0


def _tham_so_slide(meta: dict) -> dict:
    def j(k):
        v = meta.get(k)
        if isinstance(v, str):
            try:
                v = json.loads(v) if v else None
            except ValueError:
                raise Loi(400, f"Tham số “{k}” không phải JSON hợp lệ.")
        return v
    return {"tuy_chinh": j("tuy_chinh"), "tieu_de_moi": j("tieu_de_moi")}


def _slide_phan_tich(q: dict, body: bytes) -> Resp:
    meta, tep = _tach_khung(body, q)
    if not tep.get("tep") and not tep.get("pptx"):
        raise Loi(400, "Chưa nhận được file.")
    data = tep.get("pptx") or tep["tep"]
    slides, w0 = _doc_pptx(data, meta.get("ten", ""))
    try:
        prof = ap_dung_tuy_chinh(load_profile(ten_profile(meta.get("kieu", "chuan"))), _tham_so_slide(meta)["tuy_chinh"])
    except ValueError as e:
        raise Loi(400, str(e))
    kq = slide_lint.phan_tich(slides, prof)
    kq["tieu_de"] = [{"stt": i, "tieu_de": s.title, "so_y": len(s.items)} for i, s in enumerate(slides, 1)]
    kq["canh_bao"] = w0
    return _json(kq)


def _slide_chuan_hoa(q: dict, body: bytes) -> Resp:
    meta, tep = _tach_khung(body, q)
    data = tep.get("pptx") or tep.get("tep")
    if not data:
        raise Loi(400, "Chưa nhận được nội dung file.")
    slides, w0 = _doc_pptx(data, meta.get("ten", ""))
    ts = _tham_so_slide(meta)
    try:
        out, w1, n = dung_slide(slides, _thong_tin_tu_query(meta), profile=ten_profile(meta.get("kieu", "chuan")),
                                tuy_chinh=ts["tuy_chinh"], tieu_de_moi=ts["tieu_de_moi"])
    except ValueError as e:
        raise Loi(400, str(e))
    tom_tat = []
    try:                                              # kiểm tra lại chính file vừa dựng theo tiêu chuẩn của kiểu slide đã chọn
        from .importers.pptx_reader import doc_pptx
        from .slide_lint import phan_tich
        kq = phan_tich(doc_pptx(out)[0], load_profile(ten_profile(meta.get("kieu", "chuan"))))
        t = kq["tong"]
        tom_tat = [f"Kiểm tra tiêu chuẩn sau chuẩn hóa: {t['dat']} đạt, {t['canh_bao']} cảnh báo, {t['loi']} lỗi, "
                   f"{t['goi_y']} gợi ý (chủ yếu tiêu đề chưa là câu kết luận: hãy sửa từng slide ở mục tiêu đề)."]
    except Exception:
        pass
    return _file(out, PPTX_MIME, "slide-chuan.pptx", {"X-Canh-Bao": dau_header(tom_tat + w0 + w1), "X-So-Slide": str(n)})


# --------------------------------------------------------------- thông tin
def _thong_tin(q: dict) -> Resp:
    from .renderers.slide_hoc_thuat import KIEU_SLIDE
    return _json({
        "tac_gia": {"chu_ky": ban_quyen.CHU_KY, "ban_quyen": ban_quyen.BAN_QUYEN},
        "the_thuc": mo_ta_profile("van_ban"),
        "slide": mo_ta_profile("slide"),
        "mac_dinh_the_thuc": DEFAULT_PROFILE,
        "loai_van_ban": {k: (v or "Công văn") for k, v in TEN_LOAI.items()},
        "bi_danh_slide": KIEU_SLIDE,
        "tuy_chon_lam_sach": [{"ma": k, "bat": v} for k, v in ky_tu_la.MAC_DINH.items()],
    })


def _profile(pid: str) -> Resp:
    if pid not in list_profiles():
        raise Loi(404, "Không có thể thức/kiểu slide này.")
    return _json(load_profile(pid))


# --------------------------------------------------------------- lô
def _lo_them(lo: str, q: dict, body: bytes) -> Resp:
    if not body:
        raise Loi(400, "Chưa nhận được nội dung file.")
    tep_phu: dict = {}
    if khung.la_khung(body):                       # PDF kèm ảnh trang / ảnh watermark / ảnh ký
        tt, tep = _tach_khung(body, q)
        data = tep.pop("tep", None) or tep.pop("pdf", None)
        if not data:
            raise Loi(400, "Thiếu file trong gói dữ liệu.")
        tep_phu = tep
    else:
        tt, data = dict(q), body
    try:
        return _json(lo_mod.xu_ly_tep(lo, tt.get("ten", ""), data, tt, tep_phu))
    except (ValueError, FileNotFoundError) as e:
        raise Loi(400, str(e))


# --------------------------------------------------------------- điểm vào
def xu_ly(method: str, path: str, q: dict, body: bytes = b"") -> Resp:
    """Điều phối một yêu cầu; luôn trả Resp (lỗi -> JSON {detail})."""
    try:
        r = _dinh_tuyen(method.upper(), path, q, body)
    except Loi as e:
        r = _json({"detail": e.detail}, e.status)
    except Exception as e:  # lỗi bất ngờ: báo về giao diện thay vì làm treo
        r = _json({"detail": f"Lỗi xử lý: {type(e).__name__}: {e}"}, 500)
    r.headers["X-Tac-Gia"] = ban_quyen.CHU_KY
    return r


def _dinh_tuyen(m: str, p: str, q: dict, body: bytes) -> Resp:
    if m == "GET":
        if p == "/api/tac-gia":
            return _json({"chu_ky": ban_quyen.CHU_KY, "ban_quyen": ban_quyen.BAN_QUYEN})
        if p == "/api/loai-van-ban":
            pid = str(q.get("the_thuc", "")).strip()
            if pid:
                if pid not in list_profiles():
                    raise Loi(404, "Không có thể thức này.")
                from .loai_vb import loai_cua_profile
                return _json({"loai": {k: {"nhan": e["nhan"], "kinh_gui": e["kinh_gui"], "goi_y": e.get("goi_y", ""),
                                           "trich_yeu_nhan": e.get("trich_yeu_nhan", ""), "muc": [m["ten"] for m in e["muc"]]}
                                       for k, e in loai_cua_profile(load_profile(pid)).items()}})
            return _json({k: (v or "Công văn") for k, v in TEN_LOAI.items()})
        if p == "/api/thong-tin":
            return _thong_tin(q)
        x = re.fullmatch(r"/api/profile/([\w\-]+)", p)
        if x:
            return _profile(x.group(1))
        x = re.fullmatch(r"/api/lo/([^/]+)/zip", p)
        if x:
            try:
                return _file(lo_mod.tao_zip(x.group(1)).read_bytes(), "application/zip", "ket-qua-chuan-hoa.zip")
            except (ValueError, FileNotFoundError) as e:
                raise Loi(404, str(e))
    elif m == "POST":
        tuyen = {"/api/import": _import, "/api/cong-van/docx": _cong_van, "/api/ra-soat": _ra_soat, "/api/ky-tu-la": _ky_tu_la,
                 "/api/pdf/phan-tich": _pdf_phan_tich, "/api/pdf/xu-ly": _pdf_xu_ly,
                 "/api/slide/phan-tich": _slide_phan_tich, "/api/slide/chuan-hoa": _slide_chuan_hoa}
        if p in tuyen:
            return tuyen[p](q, body)
        if p == "/api/lo/tao":
            return _json({"lo": lo_mod.tao_lo()})
        x = re.fullmatch(r"/api/lo/([^/]+)/them", p)
        if x:
            return _lo_them(x.group(1), q, body)
    raise Loi(404, "Không tìm thấy đường dẫn.")
