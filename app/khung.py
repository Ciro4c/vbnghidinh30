# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Gói nhiều phần (JSON + nhiều tệp nhị phân) vào một thân yêu cầu/phản hồi.

Định dạng: b"VBN1" + 4 byte (big-endian) độ dài JSON + JSON (UTF-8) + các tệp nối liền nhau.
JSON có khóa "tep": [{"ten": "...", "co": số_byte}, ...] mô tả các tệp theo thứ tự.
Dùng chung cho FastAPI (máy cá nhân) và cầu nối Pyodide (trình duyệt), tránh phải gửi base64.
"""
from __future__ import annotations

import json
import struct

MAGIC = b"VBN1"
MIME = "application/x-vbn-khung"


class LoiKhung(ValueError):
    pass


def dong_goi(meta: dict, tep: list[tuple[str, bytes]] | None = None) -> bytes:
    meta = dict(meta)
    meta["tep"] = [{"ten": t, "co": len(b)} for t, b in (tep or [])]
    js = json.dumps(meta, ensure_ascii=False).encode("utf-8")
    return MAGIC + struct.pack(">I", len(js)) + js + b"".join(b for _, b in (tep or []))


def la_khung(data: bytes) -> bool:
    return data[:4] == MAGIC


def tach(data: bytes) -> tuple[dict, dict[str, bytes]]:
    """Trả về (meta, {tên: bytes}); các tệp trùng tên giữ bản cuối."""
    if not la_khung(data) or len(data) < 8:
        raise LoiKhung("Dữ liệu gửi lên không đúng định dạng.")
    (n,) = struct.unpack(">I", data[4:8])
    if 8 + n > len(data):
        raise LoiKhung("Dữ liệu gửi lên bị cắt cụt.")
    try:
        meta = json.loads(data[8:8 + n].decode("utf-8"))
    except ValueError as e:
        raise LoiKhung("Phần mô tả gửi lên không đọc được.") from e
    pos = 8 + n
    tep: dict[str, bytes] = {}
    for m in meta.get("tep", []):
        co = int(m.get("co", 0))
        if co < 0 or pos + co > len(data):
            raise LoiKhung("Dữ liệu gửi lên bị cắt cụt.")
        tep[str(m["ten"])] = data[pos:pos + co]
        pos += co
    return meta, tep
