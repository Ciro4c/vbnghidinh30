# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Đối chiếu đầy đủ: mọi từ trong file Word gốc phải còn trong kết quả nhập (các ô của form).

Dùng để báo ngay khi nhận diện làm rơi chữ, thay vì để người dùng tự phát hiện sau khi xuất file."""
from __future__ import annotations

import io
import re
import unicodedata
from collections import Counter

from docx import Document
from docx.oxml.ns import qn

# Chữ do hệ thống tự dựng lại theo thể thức (quốc hiệu, tiêu ngữ, nhãn "Số", "Nơi nhận"...): không tính là mất
_TU_HE_THONG = set("br rightarrow cộng hòa xã hội chủ nghĩa việt nam độc lập tự do hạnh phúc số ngày tháng năm nơi nhận kính gửi".split())
_RE_TU = re.compile(r"[^\W\d_]+(?:['’-][^\W\d_]+)*", re.U)


def _chuan(t: str) -> list[str]:
    t = unicodedata.normalize("NFC", t.lower())
    return [w for w in _RE_TU.findall(t) if w not in _TU_HE_THONG]


def _chu_nguon(data: bytes) -> list[str]:
    doc = Document(io.BytesIO(data))
    out: list[str] = []
    body = doc.element.body
    bo_qua_toc = set()
    for p in body.iter(qn("w:p")):
        ten = ""
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
            ten = (ppr.find(qn("w:pStyle")).get(qn("w:val")) or "").lower()
        if re.match(r"(toc|muclc|mụclục)\d*$", ten.replace(" ", "")):
            bo_qua_toc.add(id(p))
    for p in body.iter(qn("w:p")):
        if id(p) in bo_qua_toc:
            continue
        anc = [a for a in p.iterancestors() if a.tag == qn("w:sdt")]
        if any(a.find(qn("w:sdtPr")) is not None and any(
                "table of contents" in (g.get(qn("w:val")) or "").lower() for g in a.find(qn("w:sdtPr")).iter(qn("w:docPartGallery")))
                for a in anc):
            continue
        for el in p.iter(qn("w:t")):
            if any(a.tag in (qn("w:del"),) for a in el.iterancestors()):
                continue
            out.append(el.text or "")
        out.append(" ")
    return out


def chu_ket_qua(fields: dict) -> str:
    gop = []
    for k, v in fields.items():
        if isinstance(v, list):
            gop.extend(str(x) for x in v)
        elif isinstance(v, str):
            gop.append(v)
    t = "\n".join(gop)
    return re.sub(r"\[\[[^\]]*\]\]|\*+|[|#^<]", " ", t.replace("\\*", "*").replace("<br>", " "))


def thieu(data: bytes, fields: dict) -> list[str]:
    """Danh sách từ có trong file gốc nhưng không có (đủ số lần) trong kết quả."""
    nguon = Counter(_chuan("".join(_chu_nguon(data))))
    from ..loai_vb import tat_ca_ten
    ten_loai = tat_ca_ten().get(str(fields.get("loai_van_ban") or ""), "")      # tên loại do hệ thống tự in
    kq = Counter(_chuan(chu_ket_qua(fields) + " " + ten_loai))
    nguon.subtract(kq)
    return [w for w, n in nguon.items() for _ in range(max(n, 0))]


def canh_bao(data: bytes, fields: dict) -> list[str]:
    try:
        mat = thieu(data, fields)
    except Exception:
        return []
    if not mat:
        return []
    mau = ", ".join(dict.fromkeys(mat[:12]))
    return [f"Đối chiếu đầy đủ: có {len(mat)} từ trong file gốc không thấy trong kết quả (ví dụ: {mau}). "
            "Hãy so lại với file Word gốc trước khi dùng."]
