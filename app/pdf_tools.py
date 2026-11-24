# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Xử lý PDF: phân tích, dò ký tự lạ, kiểm tra thể thức, gắn dấu trang (bookmark), watermark chữ/hình,
lồng chữ ký và mộc, lưu dạng ảnh (flatten).

Chỉ dùng thư viện thuần Python (pypdf, Pillow) để chạy được cả trong trình duyệt (Pyodide).
Việc vẽ trang PDF thành ảnh (khi chọn "lưu dạng ảnh") do trình duyệt làm bằng pdf.js rồi gửi lên;
khi chạy ở máy cá nhân/kiểm thử có thể dùng pypdfium2 nếu đã cài (xem `ve_trang_bang_pdfium`).
"""
from __future__ import annotations

import io
import re
import zlib

from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject, EncodedStreamObject,
                           NameObject, NumberObject)

from . import anh_ky, ban_quyen, ky_tu_la
from .profile import DEFAULT_PROFILE, load_profile

PT_MM = 25.4 / 72
MAX_TRANG = 2000
VI_TRI = {   # tên -> (ngang 0..1, dọc 0..1 từ trên xuống)
    "tren_trai": (0, 0), "tren_giua": (.5, 0), "tren_phai": (1, 0),
    "giua_trai": (0, .5), "giua": (.5, .5), "giua_phai": (1, .5),
    "duoi_trai": (0, 1), "duoi_giua": (.5, 1), "duoi_phai": (1, 1),
}


class LoiPdf(ValueError):
    pass


# ------------------------------------------------------------------ mở file
def mo_pdf(data: bytes) -> PdfReader:
    try:
        r = PdfReader(io.BytesIO(data))
        if r.is_encrypted:
            try:
                ok = r.decrypt("")
            except Exception as e:                       # thiếu thư viện giải mã AES, hoặc mật khẩu khác rỗng
                raise LoiPdf("PDF đặt mật khẩu hoặc mã hóa mà ứng dụng không giải được. "
                             "Hãy mở bằng trình đọc PDF rồi In/Lưu thành PDF mới không mật khẩu.") from e
            if not ok:
                raise LoiPdf("PDF đặt mật khẩu. Hãy gỡ mật khẩu rồi tải lại.")
        n = len(r.pages)
    except LoiPdf:
        raise
    except Exception as e:
        raise LoiPdf("Không đọc được file PDF (có thể bị hỏng).") from e
    if n < 1:
        raise LoiPdf("File PDF không có trang nào.")
    if n > MAX_TRANG:
        raise LoiPdf(f"PDF quá nhiều trang (tối đa {MAX_TRANG}).")
    return r


def _hop_nhin(page) -> tuple[float, float, float, float, int]:
    """(x0, y0, rộng, cao) của khung hiển thị (CropBox) và góc xoay; rộng/cao đã tính theo chiều xoay hiển thị."""
    box = page.cropbox
    x0, y0 = float(box.left), float(box.bottom)
    w, h = float(box.width), float(box.height)
    rot = int(page.get("/Rotate", 0) or 0) % 360
    return x0, y0, w, h, rot


def kich_thuoc_hien_thi(page) -> tuple[float, float]:
    _, _, w, h, rot = _hop_nhin(page)
    return (h, w) if rot in (90, 270) else (w, h)


# ------------------------------------------------------------------ văn bản và phân tích
def trich_van_ban(reader: PdfReader) -> list[str]:
    out = []
    for p in reader.pages:
        try:
            t = p.extract_text() or ""
        except Exception:
            t = ""
        out.append(t.replace("\r", ""))
    return out


def doc_outline(reader: PdfReader) -> list[dict]:
    """Dấu trang sẵn có của PDF: [{cap, tieu_de, trang(1-based)}]."""
    out: list[dict] = []

    def di(items, cap):
        for it in items:
            if isinstance(it, list):
                di(it, cap + 1)
                continue
            try:
                tr = reader.get_destination_page_number(it) + 1
            except Exception:
                continue
            if tr >= 1:
                out.append({"cap": cap, "tieu_de": str(getattr(it, "title", "") or "").strip() or "(không tên)",
                            "trang": tr})
    try:
        di(reader.outline, 1)
    except Exception:
        return []
    return out


def phat_hien_ky_tu_la(trang_van_ban: list[str]) -> list[dict]:
    ds = []
    nhieu = ("khoang_trang_thua", "dau_cau_sai", "ngoac_kep_lan")       # do cách pypdf ghép chữ, không phải lỗi thật
    for i, t in enumerate(trang_van_ban, 1):
        if t.strip():
            ds.extend(p.to_dict() for p in ky_tu_la.quet(t, trang=i, toi_da=80) if p.ma not in nhieu)
    return ds[:400]


def goi_y_bookmark(trang_van_ban: list[str], prof: dict | None = None) -> list[dict]:
    """Dò các dòng tiêu đề theo `bookmark.mau` của profile (quốc hiệu, tên loại, Chương, Mục, Điều, Phụ lục, Nơi nhận...)."""
    prof = prof or load_profile(DEFAULT_PROFILE)
    mau = []
    for m in prof.get("bookmark", {}).get("mau", []):
        try:
            mau.append((int(m["cap"]), m.get("ten", ""), re.compile(m["mau"])))
        except (KeyError, re.error, ValueError):
            continue
    out: list[dict] = []
    thay = set()
    for i, t in enumerate(trang_van_ban, 1):
        dong = [x.strip() for x in t.split("\n") if x.strip()]
        for k, ln in enumerate(dong):
            ln_g = re.sub(r"\s+", " ", ln)
            for cap, ten, rx in mau:
                if rx.search(ln_g):
                    nhan = ten or ln_g
                    if len(nhan) > 90:
                        nhan = nhan[:87].rstrip() + "…"
                    # tên loại + trích yếu ở dòng sau
                    if not ten and ln_g.isupper() and k + 1 < len(dong) and cap == 1 and i == 1 and len(ln_g) < 40 \
                            and not ln_g.startswith("PHỤ LỤC") and not dong[k + 1].isupper():
                        nhan = f"{ln_g.rstrip(':')} - {dong[k + 1][:60]}"
                    khoa = (i, nhan)
                    if khoa in thay:
                        break
                    thay.add(khoa)
                    out.append({"cap": cap, "tieu_de": nhan, "trang": i})
                    break
        if len(out) > 400:
            break
    return out


_RE_THE_THUC = {
    "quoc_hieu": (r"CỘNG\s*HÒA\s*XÃ\s*HỘI\s*CHỦ\s*NGHĨA\s*VIỆT\s*NAM|ĐẢNG\s*CỘNG\s*SẢN\s*VIỆT\s*NAM", "đầu trang 1"),
    "tieu_ngu": (r"Độc\s*lập\s*[-–]\s*Tự\s*do\s*[-–]\s*Hạnh\s*phúc", "đầu trang 1"),
    "so_ky_hieu": (r"\bSố\s*:?\s*\d+\s*[-/][^\s]+", "trang 1"),
    "dia_danh_ngay": (r"ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}", "trang 1"),
    "kinh_gui": (r"Kính\s*gửi\s*:", "trang 1"),
    "noi_nhan": (r"Nơi\s*nhận\s*:?", "trang cuối"),
}


def kiem_tra_the_thuc(trang_van_ban: list[str], prof: dict, thanh_phan: dict | None = None,
                      kich_thuoc: tuple[float, float] | None = None) -> list[dict]:
    """Danh sách kiểm tra: mỗi thành phần thể thức đang được chọn có xuất hiện trong văn bản không.
    Thành phần bị bỏ chọn thì không bị coi là thiếu. Kết quả chỉ mang tính gợi ý (dò theo chữ)."""
    tp = dict(prof.get("thanh_phan", {}))
    tp.update({k: bool(v) for k, v in (thanh_phan or {}).items() if k in tp})
    nhan = prof.get("nhan_thanh_phan", {})
    dau = "\n".join(trang_van_ban[:1])
    cuoi = "\n".join(trang_van_ban[-2:]) if trang_van_ban else ""
    toan = "\n".join(trang_van_ban)
    out = []
    if not toan.strip():
        return [{"thanh_phan": "van_ban", "nhan": "Văn bản", "ket_qua": "khong_ro",
                 "ghi_chu": "PDF không có lớp chữ (dạng ảnh/scan) nên không kiểm tra được theo chữ."}]
    for k, (rx, vung) in _RE_THE_THUC.items():
        if not tp.get(k, True):
            continue
        nguon = cuoi if k == "noi_nhan" else dau
        if k == "quoc_hieu":
            qh = prof.get("header", {}).get("quoc_hieu_text", "")
            rx = re.escape(qh).replace(r"\ ", r"\s*") if qh else rx
        ok = bool(re.search(rx, nguon, re.I if k not in ("quoc_hieu",) else 0))
        out.append({"thanh_phan": k, "nhan": nhan.get(k, k), "ket_qua": "co" if ok else "thieu",
                    "ghi_chu": "" if ok else f"Không thấy ở {vung}."})
    if tp.get("co_quan", True):
        dong = [x.strip() for x in dau.split("\n") if x.strip()][:10]
        ok = any(x.isupper() and len(x) >= 5 and not re.search(r"CỘNG HÒA|ĐẢNG CỘNG SẢN", x) for x in dong)
        out.append({"thanh_phan": "co_quan", "nhan": nhan.get("co_quan", "Cơ quan"), "ket_qua": "co" if ok else "thieu",
                    "ghi_chu": "" if ok else "Không thấy dòng chữ in hoa nào là tên cơ quan ở đầu trang 1."})
    if tp.get("so_trang", True) and len(trang_van_ban) >= 2:
        p2 = [x.strip() for x in trang_van_ban[1].split("\n") if x.strip()]
        ok = bool(p2) and (p2[0] == "2" or p2[-1] == "2")
        out.append({"thanh_phan": "so_trang", "nhan": nhan.get("so_trang", "Số trang"),
                    "ket_qua": "co" if ok else "khong_ro",
                    "ghi_chu": "" if ok else "Không dò được số “2” ở đầu/cuối trang 2 (có thể số trang là hình/trường đặc biệt)."})
    if kich_thuoc:
        w, h = kich_thuoc[0] * PT_MM, kich_thuoc[1] * PT_MM
        pg = prof.get("page", {"width_mm": 210, "height_mm": 297})
        ok = abs(w - pg["width_mm"]) < 3 and abs(h - pg["height_mm"]) < 3
        out.append({"thanh_phan": "kho_giay", "nhan": "Khổ giấy", "ket_qua": "co" if ok else "thieu",
                    "ghi_chu": f"{w:.0f} x {h:.0f} mm" + ("" if ok else f" (thể thức yêu cầu {pg['width_mm']} x {pg['height_mm']} mm)")})
    return out


def uoc_luong_le(reader: PdfReader, trang: int = 0) -> dict | None:
    """Ước lượng lề (mm) của một trang dựa trên vị trí các dòng chữ. Chỉ là ước lượng."""
    try:
        page = reader.pages[trang]
        x0, y0, w, h, rot = _hop_nhin(page)
        if rot:
            return None
        xs: list[float] = []
        ys: list[float] = []

        def visitor(text, cm, tm, font_dict, font_size):
            if not text.strip():
                return
            x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
            y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
            xs.append(x - x0)
            ys.append(y - y0)
            xs.append(x - x0 + len(text.strip()) * font_size * 0.45)
        page.extract_text(visitor_text=visitor)
        if len(xs) < 4:
            return None
        return {"trai_mm": round(min(xs) * PT_MM, 1), "phai_mm": round((w - max(xs)) * PT_MM, 1),
                "tren_mm": round((h - max(ys)) * PT_MM, 1)}
    except Exception:
        return None


def phan_tich(data: bytes, profile: str = DEFAULT_PROFILE, thanh_phan: dict | None = None) -> dict:
    """Phân tích một PDF: thông tin chung, ký tự lạ, kiểm tra thể thức, dấu trang gợi ý và dấu trang sẵn có."""
    r = mo_pdf(data)
    prof = load_profile(profile)
    van_ban = trich_van_ban(r)
    co_chu = sum(len(t.strip()) for t in van_ban[:5]) > 30
    kt = kich_thuoc_hien_thi(r.pages[0])
    meta = r.metadata or {}
    out = {
        "so_trang": len(r.pages),
        "kich_thuoc_mm": [round(kt[0] * PT_MM, 1), round(kt[1] * PT_MM, 1)],
        "co_lop_chu": co_chu,
        "tieu_de": str(meta.get("/Title") or ""),
        "bookmark_goc": doc_outline(r),
        "bookmark_goi_y": goi_y_bookmark(van_ban, prof) if co_chu else [],
        "ky_tu_la": phat_hien_ky_tu_la(van_ban) if co_chu else [],
        "kiem_tra_the_thuc": kiem_tra_the_thuc(van_ban, prof, thanh_phan, kt),
        "le_trang_1": uoc_luong_le(r, 0) if co_chu else None,
        "canh_bao": [],
    }
    if not co_chu:
        out["canh_bao"].append("PDF không có lớp chữ (dạng ảnh/scan): không dò được ký tự lạ, thể thức, dấu trang theo chữ. "
                               "Vẫn có thể gắn watermark, lồng chữ ký/mộc và gắn dấu trang thủ công.")
    out["tom_tat_ky_tu_la"] = ky_tu_la.tom_tat([ky_tu_la.PhatHien(**{k: v for k, v in d.items()}) for d in out["ky_tu_la"]])
    return out


# ------------------------------------------------------------------ chọn trang
def chon_trang(dac_ta, n: int) -> set[int]:
    """'tat_ca', 'dau', 'cuoi', 'le', 'chan', '1,3,5-8', '2-cuoi'. Trả về tập chỉ số trang 0-based."""
    s = str(dac_ta if dac_ta not in (None, "") else "tat_ca").strip().lower()
    if s in ("tat_ca", "all", "*"):
        return set(range(n))
    out: set[int] = set()
    for tok in re.split(r"[,;\s]+", s):
        if not tok:
            continue
        if tok in ("dau", "first"):
            out.add(0)
        elif tok in ("cuoi", "last"):
            out.add(n - 1)
        elif tok in ("le", "odd"):
            out |= set(range(0, n, 2))                # trang 1, 3, 5... (chỉ số 0, 2, 4...)
        elif tok in ("chan", "even"):
            out |= set(range(1, n, 2))
        else:
            m = re.fullmatch(r"(\d+|cuoi|last)(?:-(\d+|cuoi|last))?", tok)
            if not m:
                raise LoiPdf(f"Không hiểu cách chọn trang “{tok}”. Dùng: tat_ca, dau, cuoi, le, chan, hoặc 1,3,5-8.")
            a = n if m.group(1) in ("cuoi", "last") else int(m.group(1))
            b = a if m.group(2) is None else (n if m.group(2) in ("cuoi", "last") else int(m.group(2)))
            if a > b:
                a, b = b, a
            out |= {i - 1 for i in range(max(a, 1), min(b, n) + 1)}
    return out


# ------------------------------------------------------------------ dựng lớp watermark/chữ ký/mộc (ảnh)
def _tao_item_watermark(cfg: dict, loai: str, tep: dict) -> Image.Image:
    if loai == "chu":
        return anh_ky.ve_chu(str(cfg.get("chu", "")), co_px=220, mau=str(cfg.get("mau", "808080")),
                             dam=bool(cfg.get("dam", True)), kieu=str(cfg.get("font", "sans")))
    if "wm_anh" not in tep:
        raise LoiPdf("Chưa chọn ảnh watermark.")
    try:
        return anh_ky.chuan_bi(tep["wm_anh"], tach_nen=bool(cfg.get("tach_nen", False)), do_dam=100, cat=False)
    except anh_ky.LoiAnh as e:
        raise LoiPdf(str(e)) from e


def dung_lop_watermark(cfg: dict, loai: str, W: int, H: int, px_mm: float, tep: dict) -> Image.Image:
    """Lớp RGBA kích thước WxH px chứa watermark (một hình ở giữa, lát lặp, hoặc ở vị trí chọn)."""
    item = _tao_item_watermark(cfg, loai, tep)
    kieu = str(cfg.get("kieu", "giua"))
    goc = anh_ky.goc_xoay_hop_le(cfg.get("goc", 45 if (kieu in ("giua", "lat") and loai == "chu") else 0))
    mac_dinh_rong = {"giua": 60, "lat": 28, "vi_tri": 30}.get(kieu, 60)
    rong_pct = max(3.0, min(100.0, float(cfg.get("rong_pct", mac_dinh_rong))))
    it = anh_ky.xoay(item, goc)
    it = anh_ky.vua_khung(it, W * rong_pct / 100, H * 0.95)
    it = anh_ky.chinh_do_dam(it, float(cfg.get("do_dam", 25)))
    lop = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    if kieu == "lat":
        gap = max(0.0, float(cfg.get("khoang_cach_pct", 80))) / 100
        return anh_ky.gach_lat(it, W, H, max(1, int(it.width * (1 + gap))), max(1, int(it.height * (1 + gap))))
    if kieu == "vi_tri":
        ax, ay = VI_TRI.get(str(cfg.get("vi_tri", "giua")), (.5, .5))
        le = anh_ky.mm_sang_px(float(cfg.get("le_mm", 10)), px_mm * 25.4)
        x = int(le + ax * (W - 2 * le - it.width))
        y = int(le + ay * (H - 2 * le - it.height))
    else:
        x, y = (W - it.width) // 2, (H - it.height) // 2
    anh_ky._dan_cat(lop, it, x, y)
    return lop


def chuan_bi_dau(cfg: dict, khoa_tep: str, tep: dict, rong_mac_dinh_mm: float) -> tuple[Image.Image, float] | None:
    """(ảnh chữ ký/mộc đã xử lý, bề rộng mm) hoặc None nếu không có ảnh."""
    if khoa_tep not in tep:
        return None
    try:
        im = anh_ky.chuan_bi(tep[khoa_tep], tach_nen=bool(cfg.get("tach_nen", True)),
                             do_dam=float(cfg.get("do_dam", 100)), nhuom=str(cfg.get("nhuom", "")))
    except anh_ky.LoiAnh as e:
        raise LoiPdf(str(e)) from e
    w = float(cfg.get("duong_kinh_mm") or cfg.get("rong_mm") or rong_mac_dinh_mm)
    return im, max(5.0, min(w, 300.0))


def vi_tri_dau(cfg: dict, w_mm: float, h_mm: float, trang_w_mm: float, trang_h_mm: float) -> tuple[float, float]:
    """Góc trên-trái (mm, tính từ góc trên-trái trang) của chữ ký/mộc, theo tâm x_pct, y_pct."""
    x = float(cfg.get("x_pct", 75)) / 100 * trang_w_mm - w_mm / 2
    y = float(cfg.get("y_pct", 82)) / 100 * trang_h_mm - h_mm / 2
    return (max(0.0, min(x, trang_w_mm - w_mm)), max(0.0, min(y, trang_h_mm - h_mm)))


# ------------------------------------------------------------------ ghi đối tượng PDF
def _luong_nho(rieng: bytes, **kv) -> EncodedStreamObject:
    s = EncodedStreamObject()
    s._data = rieng
    for k, v in kv.items():
        s[NameObject("/" + k)] = v
    return s


def _xobject_rgba(writer: PdfWriter, im: Image.Image):
    """Ảnh RGBA -> XObject ảnh (RGB nén Flate + mặt nạ alpha SMask)."""
    im = im.convert("RGBA")
    w, h = im.size
    rgb = zlib.compress(im.convert("RGB").tobytes(), 6)
    alpha = im.getchannel("A")
    chung = {"Type": NameObject("/XObject"), "Subtype": NameObject("/Image"), "Width": NumberObject(w),
             "Height": NumberObject(h), "BitsPerComponent": NumberObject(8), "Filter": NameObject("/FlateDecode")}
    s = _luong_nho(rgb, ColorSpace=NameObject("/DeviceRGB"), **chung)
    lo, _ = alpha.getextrema()
    if lo < 255:
        sm = _luong_nho(zlib.compress(alpha.tobytes(), 6), ColorSpace=NameObject("/DeviceGray"), **chung)
        s[NameObject("/SMask")] = writer._add_object(sm)
    return writer._add_object(s)


def _xobject_jpeg(writer: PdfWriter, jpeg: bytes, w: int, h: int, xam: bool):
    s = _luong_nho(jpeg, Type=NameObject("/XObject"), Subtype=NameObject("/Image"), Width=NumberObject(w),
                   Height=NumberObject(h), BitsPerComponent=NumberObject(8), Filter=NameObject("/DCTDecode"),
                   ColorSpace=NameObject("/DeviceGray" if xam else "/DeviceRGB"))
    return writer._add_object(s)


def _luong(writer: PdfWriter, b: bytes):
    s = DecodedStreamObject()
    s.set_data(b)
    return writer._add_object(s)


def _ma_tran_xoay(x0, y0, W, H, rot) -> tuple[float, ...]:
    """Ma trận cm đưa hệ tọa độ "hiển thị" (gốc dưới-trái, đã xoay) về hệ tọa độ nội dung của trang chưa xoay.
    W, H = rộng/cao khung nội dung (chưa xoay)."""
    if rot == 90:
        return (0, 1, -1, 0, x0 + W, y0)
    if rot == 180:
        return (-1, 0, 0, -1, x0 + W, y0 + H)
    if rot == 270:
        return (0, -1, 1, 0, x0, y0 + H)
    return (1, 0, 0, 1, x0, y0)


def _them_xobject_vao_trang(writer: PdfWriter, page, ten: str, ref, ops: list[bytes], ban_dau: bool = False):
    """Đăng ký XObject `ten` cho trang (tạo tài nguyên nếu thiếu)."""
    res = page.get("/Resources")
    res = res.get_object() if res is not None else None
    if res is None:
        res = DictionaryObject()
        page[NameObject("/Resources")] = res
    xo = res.get("/XObject")
    xo = xo.get_object() if xo is not None else None
    if xo is None:
        xo = DictionaryObject()
        res[NameObject("/XObject")] = xo
    xo[NameObject("/" + ten)] = ref


def _them_noi_dung(writer: PdfWriter, page, ops: bytes):
    """Chèn thêm lớp vẽ SAU nội dung gốc; nội dung gốc được bọc q ... Q để không rò trạng thái đồ họa."""
    cur = page.get("/Contents")
    arr = ArrayObject()
    arr.append(_luong(writer, b"q\n"))
    if cur is not None:
        obj = cur.get_object()
        if isinstance(obj, ArrayObject):
            arr.extend(obj)
        elif hasattr(cur, "idnum"):
            arr.append(cur)
        else:
            arr.append(writer._add_object(obj))
    arr.append(_luong(writer, b"\nQ\n"))
    arr.append(_luong(writer, ops))
    page[NameObject("/Contents")] = arr


def _lenh_ve_anh(ten: str, u: float, v: float, w: float, h: float, mt: tuple[float, ...]) -> bytes:
    a, b, c, d, e, f = mt
    return (f"q {a:g} {b:g} {c:g} {d:g} {e:.3f} {f:.3f} cm {w:.3f} 0 0 {h:.3f} {u:.3f} {v:.3f} cm "
            f"/{ten} Do Q\n").encode("ascii")


# ------------------------------------------------------------------ xử lý chính
def _tuy_chon_mac_dinh(o: dict) -> dict:
    o = dict(o or {})
    o.setdefault("dang_anh", False)
    o.setdefault("dpi", 200)
    o.setdefault("chat_luong", 85)
    o.setdefault("den_trang", False)
    return o


def _bat(cfg) -> bool:
    return bool(cfg) and bool(cfg.get("bat", False))


def ve_trang_bang_pdfium(data: bytes, dpi: int = 200, chat_luong: int = 85, den_trang: bool = False) -> list[bytes]:
    """Dự phòng khi không có ảnh từ trình duyệt: vẽ trang bằng pypdfium2 (nếu đã cài). Trả về danh sách JPEG."""
    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise LoiPdf("Chưa có ảnh các trang (cần trình duyệt có pdf.js) và máy chưa cài pypdfium2 để tự vẽ.") from e
    pdf = pdfium.PdfDocument(data)
    out = []
    for i in range(len(pdf)):
        im = pdf[i].render(scale=dpi / 72).to_pil().convert("L" if den_trang else "RGB")
        b = io.BytesIO()
        im.save(b, "JPEG", quality=int(chat_luong), optimize=True)
        out.append(b.getvalue())
    return out


def xu_ly_pdf(data: bytes, tuy_chon: dict, tep: dict[str, bytes] | None = None,
              profile: str = DEFAULT_PROFILE) -> tuple[bytes, list[str], dict]:
    """Xử lý PDF theo `tuy_chon`; `tep` chứa các ảnh đi kèm: "trang:1".."trang:N" (JPEG/PNG đã vẽ sẵn, khi dang_anh),
    "wm_anh", "ky_anh", "dau_anh". Trả về (PDF mới, cảnh báo, thông tin)."""
    o = _tuy_chon_mac_dinh(tuy_chon)
    tep = tep or {}
    reader = mo_pdf(data)
    n = len(reader.pages)
    canh_bao: list[str] = []
    prof = load_profile(profile)

    # ---- danh sách dấu trang cuối cùng
    bm_cfg = o.get("bookmark") or {}
    van_ban = None
    ds_bm: list[dict] | None = None
    if isinstance(bm_cfg.get("danh_sach"), list):
        ds_bm = [{"cap": int(x.get("cap", 1)), "tieu_de": str(x.get("tieu_de", "")).strip(), "trang": int(x.get("trang", 1))}
                 for x in bm_cfg["danh_sach"] if str(x.get("tieu_de", "")).strip()]
    else:
        ds_bm = []
        if bm_cfg.get("giu_goc", True):
            ds_bm += doc_outline(reader)
        if bm_cfg.get("tu_dong"):
            van_ban = trich_van_ban(reader)
            ds_bm += goi_y_bookmark(van_ban, load_profile(bm_cfg.get("profile") or profile))
            ds_bm.sort(key=lambda x: (x["trang"],))
    ds_bm = [x for x in ds_bm if 1 <= x["trang"] <= n]

    wm_chu, wm_anh = o.get("watermark_chu"), o.get("watermark_anh")
    cfg_ky, cfg_dau = o.get("chu_ky"), o.get("con_dau")
    sel = {k: chon_trang((c or {}).get("trang", "tat_ca"), n) for k, c in
           (("wm_chu", wm_chu), ("wm_anh", wm_anh))}
    sel["ky"] = chon_trang((cfg_ky or {}).get("trang", "cuoi"), n)
    sel["dau"] = chon_trang((cfg_dau or {}).get("trang", "cuoi"), n)
    if _bat(wm_chu) and not str(wm_chu.get("chu", "")).strip():
        raise LoiPdf("Chưa nhập chữ watermark.")
    ky = chuan_bi_dau(cfg_ky, "ky_anh", tep, 38) if _bat(cfg_ky) else None
    dau = chuan_bi_dau(cfg_dau, "dau_anh", tep, 40) if _bat(cfg_dau) else None
    if _bat(cfg_ky) and ky is None:
        canh_bao.append("Đã bật lồng chữ ký nhưng chưa chọn ảnh chữ ký: bỏ qua.")
    if _bat(cfg_dau) and dau is None:
        canh_bao.append("Đã bật lồng con dấu nhưng chưa chọn ảnh con dấu: bỏ qua.")

    writer = PdfWriter(clone_from=reader) if not o["dang_anh"] else PdfWriter()
    lop_cache: dict = {}

    def lop_wm(loai, cfg, Wv, Hv, px_per_pt, ten):
        khoa = (loai, round(Wv), round(Hv), px_per_pt)
        if khoa not in lop_cache:
            W, H = max(2, int(Wv * px_per_pt)), max(2, int(Hv * px_per_pt))
            lop_cache[khoa] = dung_lop_watermark(cfg, loai, W, H, px_per_pt * 72 / 25.4, tep)
        return lop_cache[khoa]

    def bo_dau_trang(i, Wv, Hv):
        """[(ảnh, x_mm, y_mm, w_mm, h_mm)] chữ ký/mộc của trang i (tọa độ mm từ góc trên-trái hiển thị)."""
        out = []
        for k, (it, cfg) in (("ky", (ky, cfg_ky)), ("dau", (dau, cfg_dau))):
            if it is None or i not in sel[k]:
                continue
            im, w_mm = it
            h_mm = w_mm * im.height / im.width
            x, y = vi_tri_dau(cfg, w_mm, h_mm, Wv * PT_MM, Hv * PT_MM)
            out.append((im, x, y, w_mm, h_mm))
        return out

    # -------- dạng ảnh: dựng lại từng trang từ ảnh đã vẽ
    if o["dang_anh"]:
        anh_bytes = [tep.get(f"trang:{i + 1}") for i in range(n)]
        if any(b is None for b in anh_bytes):
            anh_bytes = ve_trang_bang_pdfium(data, int(o["dpi"]), int(o["chat_luong"]), bool(o["den_trang"]))
        for i in range(n):
            Wv, Hv = kich_thuoc_hien_thi(reader.pages[i])
            try:
                base = Image.open(io.BytesIO(anh_bytes[i]))
                base.load()
            except Exception as e:
                raise LoiPdf(f"Ảnh của trang {i + 1} không đọc được.") from e
            lop_can = []
            px_mm = base.width / (Wv * PT_MM)
            if _bat(wm_chu) and i in sel["wm_chu"]:
                lop_can.append(dung_lop_watermark(wm_chu, "chu", base.width, base.height, px_mm, tep))
            if _bat(wm_anh) and i in sel["wm_anh"]:
                lop_can.append(dung_lop_watermark(wm_anh, "anh", base.width, base.height, px_mm, tep))
            dau_trang = bo_dau_trang(i, Wv, Hv)
            gia_tri_jpeg = anh_bytes[i] if (base.format == "JPEG" and not lop_can and not dau_trang
                                            and not o["den_trang"]) else None
            if gia_tri_jpeg is None:
                im = base.convert("RGBA")
                for lop in lop_can:
                    im.alpha_composite(lop)
                for it, x, y, w_mm, h_mm in dau_trang:
                    pw = max(1, int(w_mm * px_mm))
                    ph = max(1, int(h_mm * px_mm))
                    im.alpha_composite(it.resize((pw, ph), Image.LANCZOS), (int(x * px_mm), int(y * px_mm)))
                im = im.convert("L" if o["den_trang"] else "RGB")
                b = io.BytesIO()
                im.save(b, "JPEG", quality=int(o["chat_luong"]), optimize=True)
                gia_tri_jpeg, xam = b.getvalue(), o["den_trang"]
                w_px, h_px = im.size
            else:
                w_px, h_px = base.size
                xam = base.mode == "L"
            page = writer.add_blank_page(Wv, Hv)
            ref = _xobject_jpeg(writer, gia_tri_jpeg, w_px, h_px, xam)
            _them_xobject_vao_trang(writer, page, "VbnPg", ref, [])
            page[NameObject("/Contents")] = _luong(writer, f"q {Wv:.3f} 0 0 {Hv:.3f} 0 0 cm /VbnPg Do Q\n".encode())
    else:
        # -------- giữ vector: chèn lớp watermark, chữ ký, mộc lên từng trang
        for i in range(n):
            page = writer.pages[i]
            x0, y0, W, H, rot = _hop_nhin(page)
            Wv, Hv = (H, W) if rot in (90, 270) else (W, H)
            mt = _ma_tran_xoay(x0, y0, W, H, rot)
            ops = b""
            s_px = min(150 / 72, (14_000_000 / max(Wv * Hv, 1)) ** 0.5)
            for tag, loai, cfg in (("wm_chu", "chu", wm_chu), ("wm_anh", "anh", wm_anh)):
                if _bat(cfg) and i in sel[tag]:
                    lop = lop_wm(loai, cfg, Wv, Hv, s_px, tag)
                    ref = _xobject_rgba(writer, lop)
                    nm = f"Vbn{tag}{i}"
                    _them_xobject_vao_trang(writer, page, nm, ref, [])
                    ops += _lenh_ve_anh(nm, 0, 0, Wv, Hv, mt)
            for j, (im, x, y, w_mm, h_mm) in enumerate(bo_dau_trang(i, Wv, Hv)):
                ref = _xobject_rgba(writer, im)
                nm = f"VbnDau{i}_{j}"
                _them_xobject_vao_trang(writer, page, nm, ref, [])
                w_pt, h_pt = w_mm / PT_MM, h_mm / PT_MM
                ops += _lenh_ve_anh(nm, x / PT_MM, Hv - y / PT_MM - h_pt, w_pt, h_pt, mt)
            if ops:
                _them_noi_dung(writer, page, ops)

    # -------- dấu trang
    if o["dang_anh"] or ds_bm is not None:
        _ghi_bookmark(writer, ds_bm or [], n, bo_cu=not o["dang_anh"])
        if ds_bm:
            writer.page_mode = "/UseOutlines"

    # -------- thông tin tài liệu
    meta_goc = reader.metadata or {}
    info = {"/Producer": ban_quyen.GHI_CHU}
    tt = o.get("thong_tin") or {}
    tieu_de = str(tt.get("tieu_de") or meta_goc.get("/Title") or "")
    if tieu_de:
        info["/Title"] = tieu_de
    if tt.get("tac_gia") or meta_goc.get("/Author"):
        info["/Author"] = str(tt.get("tac_gia") or meta_goc.get("/Author"))
    writer.add_metadata(info)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue(), canh_bao, {"so_trang": n, "so_bookmark": len(ds_bm or []), "dang_anh": bool(o["dang_anh"])}


def _ghi_bookmark(writer: PdfWriter, ds: list[dict], n: int, bo_cu: bool):
    """Ghi lại toàn bộ dấu trang từ danh sách (cấp lồng nhau theo `cap`)."""
    if bo_cu and "/Outlines" in writer._root_object:
        del writer._root_object[NameObject("/Outlines")]
    cha: dict[int, object] = {}
    cap_truoc = 0
    for x in ds:
        cap = max(1, min(int(x["cap"]), cap_truoc + 1, 6))
        parent = cha.get(cap - 1) if cap > 1 else None
        ref = writer.add_outline_item(x["tieu_de"][:200], max(0, min(n - 1, int(x["trang"]) - 1)), parent=parent)
        cha[cap] = ref
        for k in [k for k in cha if k > cap]:
            del cha[k]
        cap_truoc = cap


def thong_tin_trang(data: bytes) -> dict:
    r = mo_pdf(data)
    kt = kich_thuoc_hien_thi(r.pages[0])
    return {"so_trang": len(r.pages), "kich_thuoc_mm": [round(kt[0] * PT_MM, 1), round(kt[1] * PT_MM, 1)]}
