# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Xử lý ảnh chữ ký, con dấu (mộc) và watermark dùng chung cho Word và PDF.

- Tách nền trắng của ảnh chụp/scan chữ ký, mộc thành nền trong suốt (có thể tắt).
- Chỉnh độ đậm nhạt (độ đục) 0-100%.
- Nhuộm màu (vd chữ ký số màu xanh) cho ảnh dạng mực.
- Vẽ chữ watermark (hỗ trợ tiếng Việt, dùng font Liberation kèm theo trong app/fonts).
"""
from __future__ import annotations

import base64
import io
import math
from functools import lru_cache
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

FONT_DIR = Path(__file__).parent / "fonts"
_FONT = {
    ("sans", False): "LiberationSans-Regular.ttf", ("sans", True): "LiberationSans-Bold.ttf",
    ("serif", False): "LiberationSerif-Regular.ttf", ("serif", True): "LiberationSerif-Bold.ttf",
}
MAX_CANH_PX = 3000          # ảnh chữ ký/mộc lớn hơn thì thu nhỏ để file không phình to
Image.MAX_IMAGE_PIXELS = 120_000_000


class LoiAnh(ValueError):
    pass


def _bytes(nguon) -> bytes:
    if isinstance(nguon, (bytes, bytearray)):
        return bytes(nguon)
    if isinstance(nguon, dict):
        nguon = nguon.get("b64", "")
    if isinstance(nguon, str):
        s = nguon.split(",", 1)[1] if nguon.startswith("data:") else nguon
        try:
            return base64.b64decode(s, validate=False)
        except Exception as e:
            raise LoiAnh("Ảnh không hợp lệ (base64 hỏng).") from e
    raise LoiAnh("Không đọc được ảnh.")


def mo_anh(nguon) -> Image.Image:
    """Mở ảnh từ bytes / base64 / dict {"b64": ...} thành RGBA, thu nhỏ nếu quá lớn."""
    data = _bytes(nguon)
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception as e:
        raise LoiAnh("Không đọc được ảnh (chỉ nhận PNG, JPG, GIF, BMP, WebP).") from e
    im = im.convert("RGBA")
    if max(im.size) > MAX_CANH_PX:
        im.thumbnail((MAX_CANH_PX, MAX_CANH_PX), Image.LANCZOS)
    return im


def co_nen_trong_suot(im: Image.Image) -> bool:
    """Ảnh đã có phần trong suốt thật (không phải ảnh phẳng toàn đục)."""
    h = im.getchannel("A").histogram()
    return sum(h[:250]) > 0.02 * (im.width * im.height)


def tach_nen_trang(im: Image.Image, nguong: int = 235, mem: int = 35) -> Image.Image:
    """Điểm càng trắng càng trong suốt: min(R,G,B) >= nguong -> trong suốt hẳn; <= nguong-mem -> đục hẳn."""
    nguong = max(1, min(255, int(nguong)))
    mem = max(1, int(mem))
    from PIL import ImageChops
    r, g, b, a = im.split()
    trang = _kenh_nho_nhat(r, g, b)                # độ trắng của điểm = kênh nhỏ nhất
    lut = []
    for v in range(256):
        if v >= nguong:
            lut.append(0)
        elif v <= nguong - mem:
            lut.append(255)
        else:
            lut.append(int(255 * (nguong - v) / mem))
    alpha_moi = trang.point(lut)
    return Image.merge("RGBA", (r, g, b, ImageChops.multiply(alpha_moi, a)))


def _kenh_nho_nhat(r, g, b):
    from PIL import ImageChops
    return ImageChops.darker(ImageChops.darker(r, g), b)


def cat_vien(im: Image.Image, le: int = 2) -> Image.Image:
    """Cắt phần viền trong suốt xung quanh."""
    bb = im.getchannel("A").point(lambda v: 255 if v > 10 else 0).getbbox()
    if not bb:
        return im
    x0, y0, x1, y1 = bb
    return im.crop((max(0, x0 - le), max(0, y0 - le), min(im.width, x1 + le), min(im.height, y1 + le)))


def chinh_do_dam(im: Image.Image, phan_tram: float) -> Image.Image:
    """Nhân kênh alpha với phan_tram/100 (độ đục)."""
    p = max(0.0, min(100.0, float(phan_tram))) / 100
    if p >= 0.999:
        return im
    r, g, b, a = im.split()
    return Image.merge("RGBA", (r, g, b, a.point(lambda v: int(v * p))))


def nhuom_muc(im: Image.Image, mau_hex: str) -> Image.Image:
    """Đổi mọi điểm ảnh thành màu `mau_hex`, độ đục theo độ tối của điểm gốc (dùng cho chữ ký dạng mực)."""
    mau = tuple(int(mau_hex[i:i + 2], 16) for i in (0, 2, 4))
    r, g, b, a = im.split()
    do_toi = Image.merge("RGB", (r, g, b)).convert("L").point(lambda v: 255 - v)
    from PIL import ImageChops
    alpha = ImageChops.multiply(do_toi, a)
    # kéo giãn để nét mảnh vẫn đủ đậm
    alpha = alpha.point(lambda v: min(255, int(v * 1.25)))
    nen = Image.new("RGBA", im.size, mau + (0,))
    nen.putalpha(alpha)
    return nen


def chuan_bi(nguon, tach_nen: bool = True, do_dam: float = 100, nhuom: str = "",
             nguong: int = 235, cat: bool = True) -> Image.Image:
    """Chuẩn bị ảnh chữ ký/mộc để chèn: tách nền, (nhuộm màu), cắt viền, chỉnh độ đậm."""
    im = mo_anh(nguon)
    if tach_nen and not co_nen_trong_suot(im):
        im = tach_nen_trang(im, nguong)
    if nhuom and len(nhuom.lstrip("#")) == 6:
        im = nhuom_muc(im, nhuom.lstrip("#"))
    if cat:
        im = cat_vien(im)
    return chinh_do_dam(im, do_dam)


def png_bytes(im: Image.Image) -> bytes:
    b = io.BytesIO()
    im.save(b, "PNG", optimize=True)
    return b.getvalue()


# ------------------------------------------------------------------ chữ watermark
@lru_cache(maxsize=16)
def _font(kieu: str, dam: bool, co_px: int):
    ten = _FONT.get((kieu if kieu in ("sans", "serif") else "sans", bool(dam)))
    try:
        return ImageFont.truetype(str(FONT_DIR / ten), co_px)
    except OSError:
        try:
            return ImageFont.load_default(co_px)          # Pillow >= 10.1: font co giãn
        except TypeError:
            return ImageFont.load_default()


def ve_chu(text: str, co_px: int = 120, mau: str = "808080", dam: bool = True, kieu: str = "sans",
           gian_dong: float = 1.15, le_px: int = 6) -> Image.Image:
    """Vẽ chữ (nhiều dòng) thành ảnh RGBA nền trong suốt, canh giữa các dòng."""
    text = (text or "").strip("\n")
    if not text.strip():
        raise LoiAnh("Chưa nhập chữ watermark.")
    co_px = max(8, min(int(co_px), 1200))
    font = _font(kieu, dam, co_px)
    tmp = ImageDraw.Draw(Image.new("L", (8, 8)))
    spacing = int(co_px * (gian_dong - 1))
    bb = tmp.multiline_textbbox((0, 0), text, font=font, spacing=spacing, align="center")
    w, h = int(math.ceil(bb[2] - bb[0])) + 2 * le_px, int(math.ceil(bb[3] - bb[1])) + 2 * le_px
    w, h = max(w, 2), max(h, 2)
    if w * h > 60_000_000:
        raise LoiAnh("Chữ watermark quá dài hoặc quá to.")
    mau = (mau or "808080").lstrip("#")
    rgb = tuple(int(mau[i:i + 2], 16) for i in (0, 2, 4)) if len(mau) == 6 else (128, 128, 128)
    im = Image.new("RGBA", (w, h), rgb + (0,))
    ImageDraw.Draw(im).multiline_text((le_px - int(bb[0]), le_px - int(bb[1])), text, font=font, fill=rgb + (255,),
                                      spacing=spacing, align="center")
    return im


def xoay(im: Image.Image, goc: float) -> Image.Image:
    """Xoay ngược chiều kim đồng hồ `goc` độ, mở rộng khung, nền trong suốt."""
    if abs(goc) < 0.01:
        return im
    return im.rotate(goc, resample=Image.BICUBIC, expand=True, fillcolor=(0, 0, 0, 0))


def vua_khung(im: Image.Image, rong_toi_da: float, cao_toi_da: float) -> Image.Image:
    """Thu/phóng giữ tỷ lệ cho vừa khung (px)."""
    k = min(rong_toi_da / im.width, cao_toi_da / im.height)
    if k <= 0:
        raise LoiAnh("Kích thước watermark không hợp lệ.")
    w, h = max(1, int(im.width * k)), max(1, int(im.height * k))
    return im.resize((w, h), Image.LANCZOS)


def gach_lat(tam: Image.Image, rong: int, cao: int, buoc_x: int, buoc_y: int, so_le: bool = True) -> Image.Image:
    """Lát `tam` thành lưới (các hàng so le) phủ khung rong x cao. Dùng cho watermark lặp."""
    nen = Image.new("RGBA", (rong, cao), (0, 0, 0, 0))
    hang = 0
    y = -buoc_y // 2
    while y < cao:
        x0 = -(buoc_x // 2) if (so_le and hang % 2) else 0
        x = x0
        while x < rong:
            _dan_cat(nen, tam, int(x), int(y))
            x += buoc_x
        y += buoc_y
        hang += 1
    return nen


def _dan_cat(nen: Image.Image, tam: Image.Image, x: int, y: int):
    """Dán `tam` vào `nen` tại (x, y) kể cả khi tràn mép (cắt phần thừa)."""
    sx, sy = max(0, -x), max(0, -y)
    ex, ey = min(tam.width, nen.width - x), min(tam.height, nen.height - y)
    if ex <= sx or ey <= sy:
        return
    nen.alpha_composite(tam.crop((sx, sy, ex, ey)), (x + sx, y + sy))


def mm_sang_px(mm: float, dpi: float) -> int:
    return max(1, int(round(mm / 25.4 * dpi)))


def goc_xoay_hop_le(g) -> float:
    try:
        g = float(g)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if not math.isfinite(g) else max(-180.0, min(180.0, g))
