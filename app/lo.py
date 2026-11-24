# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Xử lý hàng loạt: nhiều file Word/PowerPoint trong một lô, kết quả gom thành một file ZIP.

- Mỗi file xử lý độc lập: một file lỗi không làm dừng cả lô.
- Word (.docx, .txt) được nhận diện và dựng lại theo NĐ 30 hoàn toàn tự động. Thông tin thiếu
  (cơ quan ban hành, ngày) được điền chỗ giữ chỗ và ghi vào báo cáo để rà soát sau.
- PowerPoint (.pptx) được dựng lại theo chuẩn slide (mặc định: tiêu chuẩn chuyên nghiệp; có thể chọn kiểu học thuật).
- Kết quả lưu trong thư mục tạm của hệ điều hành, tự dọn sau HAN_GIU_GIAY.
"""
from __future__ import annotations

import csv
import io
import json
import re
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from datetime import date
from pathlib import Path, PurePosixPath

GOC = Path(tempfile.gettempdir()) / "vanban-nd30-lo"
HAN_GIU_GIAY = 6 * 3600
_KHOA = threading.Lock()
_RE_LO = re.compile(r"[0-9a-f]{12}")
CHO_GIU_CO_QUAN = "[CƠ QUAN BAN HÀNH]"
NHAN_LOAI = {"van_ban": "Word", "slide": "Slide", "pdf": "PDF"}
# Chế độ của lô -> các đuôi file được xử lý. Chế độ không khớp thì file bị bỏ qua (không coi là lỗi).
CHE_DO = {"van_ban": (".docx", ".txt"), "pdf": (".pdf",), "slide": (".pptx",)}
TEN_CHE_DO = {"van_ban": "Word (.docx, .txt)", "pdf": "PDF (.pdf)", "slide": "PowerPoint (.pptx)"}


class _BoQua(Exception):
    pass


# ----------------------------------------------------------------- thư mục lô
def _thu_muc(lo: str) -> Path:
    if not _RE_LO.fullmatch(str(lo)):
        raise ValueError("Mã lô không hợp lệ.")
    d = GOC / lo
    if not d.is_dir():
        raise FileNotFoundError("Không tìm thấy lô này (có thể đã hết hạn). Hãy chạy lại.")
    return d


def don_dep() -> None:
    if not GOC.exists():
        return
    for d in GOC.iterdir():
        try:
            if d.is_dir() and time.time() - d.stat().st_mtime > HAN_GIU_GIAY:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def tao_lo() -> str:
    don_dep()
    lo = uuid.uuid4().hex[:12]
    (GOC / lo / "tep").mkdir(parents=True)
    (GOC / lo / "kq.json").write_text("[]", encoding="utf-8")
    return lo


def _doc_kq(d: Path) -> list[dict]:
    try:
        return json.loads((d / "kq.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


def ket_qua(lo: str) -> list[dict]:
    return _doc_kq(_thu_muc(lo))


# ------------------------------------------------------------------ tên file
_CAM = re.compile(r'[<>:"|?*\x00-\x1f]')


def sach_duong_dan(ten: str) -> str:
    """Đường dẫn tương đối an toàn (không '..', không ổ đĩa, không ký tự cấm)."""
    parts = []
    for p in re.split(r"[\\/]+", str(ten or "")):
        p = _CAM.sub("_", p).strip(" .")
        if p and p not in (".", ".."):
            parts.append(p[:120])
    parts = parts[-8:] or ["tep"]
    return str(PurePosixPath(*parts))


def _ten_ra(thu_muc_tep: Path, rel: str, duoi_moi: str) -> str:
    p = PurePosixPath(rel)
    base = f"{p.stem}-chuan"
    ra = str(p.with_name(base + duoi_moi))
    k = 2
    while (thu_muc_tep / ra).exists():
        ra = str(p.with_name(f"{base} ({k}){duoi_moi}"))
        k += 1
    return ra


# ------------------------------------------------------------- chuyển từng loại
def _json_tt(tt: dict, khoa: str, mac_dinh=None):
    v = (tt or {}).get(khoa)
    if v in (None, ""):
        return mac_dinh
    if not isinstance(v, str):
        return v
    try:
        return json.loads(v)
    except ValueError:
        raise ValueError(f"Tham số “{khoa}” không phải JSON hợp lệ.")


def tu_dong_van_ban(data: bytes, duoi: str, tt: dict | str | None = None) -> tuple[bytes, list[str], str]:
    tt = {"profile_van_ban": tt} if isinstance(tt, str) else (tt or {})
    profile = tt.get("profile_van_ban", "")
    from .importers.docx_reader import doc_docx_ex, doc_txt
    from .importers.phan_tich import phan_tich
    from .profile import DEFAULT_PROFILE, list_profiles
    from .renderers.cong_van import cong_van_to_bytes, parse_noi_dung
    from .schemas import CongVan

    if duoi == ".txt":
        lines, w0 = doc_txt(data)
        hinh: dict = {}
    else:
        lines, w0, hinh = doc_docx_ex(data)
    if not lines:
        raise ValueError("File không có nội dung văn bản.")
    from . import ky_tu_la
    la = []
    if duoi == ".docx":
        from .api_core import doc_tho_docx
        la = ky_tu_la.quet(doc_tho_docx(data))
    da_don: dict = {}
    if str(tt.get("lam_sach", "")) in ("1", "true"):
        moi = []
        for ln in lines:
            if ln.kind == "text":
                t, tk = ky_tu_la.lam_sach_dong(ln.text, None)
                for k2, v2 in tk.items():
                    da_don[k2] = da_don.get(k2, 0) + v2
                if not t.strip():
                    continue
                ln.text = t
            moi.append(ln)
        lines = moi
    f, w1 = phan_tich(lines)
    canh_bao = list(w0) + list(w1) + ky_tu_la.tom_tat(la, da_don)
    f["thanh_phan"] = _json_tt(tt, "thanh_phan", {}) or {}
    f["tuy_chinh"] = _json_tt(tt, "tuy_chinh_van_ban", {}) or {}
    if not f["noi_dung"].strip():
        raise ValueError("Không tách được phần nội dung văn bản.")
    if not f["co_quan_ban_hanh"].strip():
        f["co_quan_ban_hanh"] = CHO_GIU_CO_QUAN
        canh_bao.append(f"Không nhận diện được cơ quan ban hành: đã điền “{CHO_GIU_CO_QUAN}”, hãy sửa trong Word.")
    if not f["ngay"]:
        f["ngay"] = date.today().isoformat()
        canh_bao.append("Không đọc được ngày ban hành: đã dùng ngày hôm nay, hãy sửa trong Word.")
    f["hinh"] = hinh
    out = cong_van_to_bytes(CongVan(**f), profile if profile in list_profiles("van_ban") else DEFAULT_PROFILE)
    toan_van = f["noi_dung"] + "\n" + f.get("phu_luc", "")        # đếm cả phụ lục
    khoi = parse_noi_dung(toan_van)
    n_bang = sum(1 for k, _ in khoi if k in ("bang", "bang_kv"))
    n_hinh = len(re.findall(r"\[\[hinh:\w+\]\]", toan_van))
    tk = ", ".join(x for x in (f"{n_bang} bảng" if n_bang else "", f"{n_hinh} hình" if n_hinh else "") if x)
    return out, canh_bao, tk


def tu_dong_slide(data: bytes, tt: dict) -> tuple[bytes, list[str], str]:
    from .importers.pptx_reader import doc_pptx
    from .renderers.slide_hoc_thuat import ThongTin, dung_slide, ten_profile

    slides, w0 = doc_pptx(data)
    if not slides:
        raise ValueError("File không có slide nào có nội dung.")
    import json as _json
    try:
        tc = _json.loads(tt["tuy_chinh_slide"]) if tt.get("tuy_chinh_slide") else None
    except ValueError:
        raise ValueError("Tham số tuỳ chỉnh slide không hợp lệ.")
    out, w1, n = dung_slide(slides, ThongTin("", tt.get("hoc_phan", "").strip(), tt.get("giang_vien", "").strip(),
                                             tt.get("don_vi", "").strip(), tt.get("nam", "").strip()),
                               profile=ten_profile(tt.get("kieu")), tuy_chinh=tc)
    w1 = [w for w in w1 if not w.startswith("Chưa nhập đủ")]   # nhắc lặp lại cho từng file thì thừa
    return out, list(w0) + w1, f"{n} slide"


def tu_dong_pdf(data: bytes, tt: dict, tep_phu: dict | None) -> tuple[bytes, list[str], str]:
    """PDF: dấu trang, watermark, chữ ký/mộc, lưu dạng ảnh (trang đã vẽ sẵn gửi kèm dạng `trang:N`)."""
    from . import pdf_tools
    from .profile import DEFAULT_PROFILE
    o = _json_tt(tt, "tuy_chon_pdf", {}) or {}
    out, canh, info = pdf_tools.xu_ly_pdf(data, o, dict(tep_phu or {}), tt.get("profile_van_ban") or DEFAULT_PROFILE)
    n = info.get("so_trang") if isinstance(info, dict) else None
    return out, list(canh), (f"{n} trang" if n else "")


def xu_ly_tep(lo: str, ten: str, data: bytes, tt: dict | None = None, tep_phu: dict | None = None) -> dict:
    """Xử lý một file của lô; luôn trả về bản ghi kết quả (kể cả khi lỗi).
    tt["che_do"] ("van_ban" | "pdf" | "slide") quyết định lô này làm gì; file khác loại bị bỏ qua."""
    d = _thu_muc(lo)
    rel = sach_duong_dan(ten)
    duoi = PurePosixPath(rel).suffix.lower()
    tt = tt or {}
    kq = {"goc": rel, "loai": "", "trang_thai": "loi", "ra": "", "canh_bao": [], "thong_ke": ""}
    che_do = tt.get("che_do") or ""
    try:
        if che_do and che_do not in CHE_DO:
            raise ValueError("Chế độ xử lý không hợp lệ.")
        if che_do and duoi not in CHE_DO[che_do] and duoi not in (".doc", ".ppt"):
            kq.update(trang_thai="bo_qua", loai=duoi.lstrip(".").upper() or "?",
                      canh_bao=[f"Bỏ qua: lô này đang ở chế độ {TEN_CHE_DO[che_do]}."])
            raise _BoQua()
        if duoi in (".docx", ".txt"):
            kq["loai"] = NHAN_LOAI["van_ban"]
            out, cb, tk = tu_dong_van_ban(data, duoi, tt)
            duoi_ra = ".docx"
        elif duoi == ".pdf":
            kq["loai"] = NHAN_LOAI["pdf"]
            out, cb, tk = tu_dong_pdf(data, tt, tep_phu)
            duoi_ra = ".pdf"
        elif duoi == ".pptx":
            kq["loai"] = NHAN_LOAI["slide"]
            out, cb, tk = tu_dong_slide(data, tt or {})
            duoi_ra = ".pptx"
        elif duoi in (".doc", ".ppt"):
            raise ValueError("File .doc/.ppt cũ chưa đọc được: hãy mở bằng Word/PowerPoint, Lưu thành "
                             ".docx/.pptx rồi thêm lại.")
        else:
            raise ValueError("Định dạng chưa hỗ trợ (chỉ nhận .docx, .txt, .pdf, .pptx).")
        with _KHOA:
            ra = _ten_ra(d / "tep", rel, duoi_ra)
            dich = d / "tep" / ra
            dich.parent.mkdir(parents=True, exist_ok=True)
            dich.write_bytes(out)
        kq.update(ra=ra, canh_bao=cb, thong_ke=tk, trang_thai="can_ra_soat" if cb else "ok")
    except _BoQua:
        pass
    except (ValueError, FileNotFoundError) as e:
        kq["canh_bao"] = [str(e)]
    except Exception:
        kq["canh_bao"] = ["Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu)."]
    with _KHOA:
        ds = _doc_kq(d)
        kq["stt"] = len(ds) + 1
        ds.append(kq)
        (d / "kq.json").write_text(json.dumps(ds, ensure_ascii=False), encoding="utf-8")
    return kq


# ------------------------------------------------------------- báo cáo + ZIP
_TEN_TRANG_THAI = {"ok": "Đạt", "can_ra_soat": "Cần rà soát", "loi": "Lỗi", "bo_qua": "Bỏ qua"}


def bao_cao_csv(ds: list[dict]) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["STT", "File gốc", "Loại", "Kết quả", "File xuất", "Nội dung đã chuyển", "Ghi chú / cảnh báo / lỗi"])
    for k in ds:
        w.writerow([k.get("stt", ""), k["goc"], k["loai"], _TEN_TRANG_THAI.get(k["trang_thai"], k["trang_thai"]),
                    k["ra"], k.get("thong_ke", ""), " | ".join(k["canh_bao"])])
    return ("\ufeff" + buf.getvalue()).encode("utf-8")      # BOM để Excel đọc đúng tiếng Việt


def tao_zip(lo: str) -> Path:
    d = _thu_muc(lo)
    ds = _doc_kq(d)
    dich = d / "ket-qua-chuan-hoa.zip"
    with zipfile.ZipFile(dich, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted((d / "tep").rglob("*")):
            if f.is_file():
                z.write(f, "ket-qua/" + f.relative_to(d / "tep").as_posix())
        z.writestr("BAO-CAO.csv", bao_cao_csv(ds))
    return dich
