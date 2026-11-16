# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Xử lý hình ảnh dùng chung cho Word và PowerPoint.

- Hình quá lớn được thu nhỏ (cạnh dài tối đa MAX_CANH px) để file kết quả không phình to.
- Hình được mã hóa base64 để mang theo form khi người dùng rà soát (tab Văn bản hành chính).
- Hình lỗi hoặc định dạng không đọc được (EMF/WMF...) trả về None để nơi gọi báo cảnh báo,
  không làm hỏng cả file.
"""
from __future__ import annotations

import base64
import hashlib
import io

MAX_CANH = 2000          # px; khổ A4 rộng 165 mm ở 300 dpi chỉ cần ~1950 px
NGUONG_BYTES = 3_000_000  # hình nhỏ hơn mức này và không quá MAX_CANH thì giữ nguyên

_MIME = {"PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif"}
EXT = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif"}


def xu_ly(blob: bytes) -> tuple[bytes, str, int, int] | None:
    """Trả về (dữ liệu, mime, rộng_px, cao_px) hoặc None nếu không đọc được."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(blob))
        im.load()
    except Exception:
        return None
    fmt = (im.format or "").upper()
    w, h = im.size
    if w < 1 or h < 1:
        return None
    if fmt in _MIME and max(w, h) <= MAX_CANH and len(blob) <= NGUONG_BYTES:
        return blob, _MIME[fmt], w, h
    try:
        if max(w, h) > MAX_CANH:
            im.thumbnail((MAX_CANH, MAX_CANH), Image.LANCZOS)
        co_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
        buf = io.BytesIO()
        if co_alpha or fmt in ("PNG", "GIF", "BMP", "TIFF"):
            im.save(buf, "PNG", optimize=True)
            if buf.tell() > 2_500_000 and not co_alpha:   # ảnh chụp lớn: JPEG gọn hơn nhiều
                buf = io.BytesIO()
                im.convert("RGB").save(buf, "JPEG", quality=85, optimize=True)
                mime = "image/jpeg"
            else:
                mime = "image/png"
        else:
            im.convert("RGB").save(buf, "JPEG", quality=85, optimize=True)
            mime = "image/jpeg"
        return buf.getvalue(), mime, im.size[0], im.size[1]
    except Exception:
        return None


class KhoHinh:
    """Gom hình của một văn bản, tự bỏ trùng (logo lặp lại nhiều lần chỉ giữ một bản)."""

    def __init__(self):
        self.hinh: dict[str, dict] = {}
        self._chi_muc: dict[str, str] = {}
        self.loi = 0

    def them(self, blob: bytes, w_mm: float = 0.0) -> str | None:
        key = hashlib.sha1(blob).hexdigest()
        if key in self._chi_muc:
            return self._chi_muc[key]
        r = xu_ly(blob)
        if r is None:
            self.loi += 1
            return None
        data, mime, pw, ph = r
        if w_mm <= 0:
            w_mm = pw * 25.4 / 96                     # không biết kích thước hiển thị: coi như 96 dpi
        id_ = str(len(self.hinh) + 1)
        self.hinh[id_] = {"mime": mime, "b64": base64.b64encode(data).decode("ascii"),
                          "w_mm": round(w_mm, 1), "w_px": pw, "h_px": ph}
        self._chi_muc[key] = id_
        return id_


def giai_ma(h: dict) -> tuple[bytes, float, float] | None:
    """Từ bản ghi hình (dict) lấy ra (dữ liệu, rộng mm, cao mm); None nếu hỏng."""
    try:
        data = base64.b64decode(h["b64"])
        pw, ph = float(h.get("w_px") or 0), float(h.get("h_px") or 0)
        if pw <= 0 or ph <= 0:
            from PIL import Image
            pw, ph = Image.open(io.BytesIO(data)).size
        w = float(h.get("w_mm") or 0) or pw * 25.4 / 96
        return data, w, w * ph / pw
    except Exception:
        return None
