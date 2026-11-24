# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Đọc file PowerPoint thô thành mô hình slide đơn giản.

Giữ lại: chữ (theo cấp), hình (thu nhỏ nếu quá lớn, sắp theo vị trí trên slide), bảng
(kể cả ô gộp: "<" gộp với ô bên trái, "^" gộp với ô phía trên), ghi chú diễn giả.
Biểu đồ được đổi thành bảng số liệu. SmartArt, video, OLE chỉ được báo để chèn lại.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from ..hinh import EXT, xu_ly


@dataclass
class SlideIn:
    title: str = ""
    items: list = field(default_factory=list)      # (cấp, văn bản, có gạch đầu dòng)
    images: list = field(default_factory=list)     # (bytes, đuôi file)
    tables: list = field(default_factory=list)     # list[list[list[str]]]
    notes: str = ""
    unsupported: list = field(default_factory=list)
    canh_bao: list = field(default_factory=list)   # cảnh báo riêng của slide (hình hỏng, biểu đồ -> bảng...)


_TITLE = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_BO_QUA = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}
_RE_GACH = re.compile(r"^[\-–—•●▪■◦○+*➢➤✓✔]\s*")


def _lam_sach(s: str) -> str:
    s = s.replace("\x0b", " ").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()


TIEU_DE_TOI_DA = 100      # tiêu đề dài hơn mức này là đoạn văn nằm nhầm trong ô tiêu đề


def _doan_chu(p) -> str:
    """Chữ của một đoạn, GIỮ chỗ xuống dòng mềm (<a:br>, Shift+Enter) thành dấu cách.
    Trước đây ghép các run nên chữ hai bên chỗ xuống dòng dính nhau (\"SprinklerĐầu\")."""
    try:
        t = p.text                      # python-pptx đổi <a:br> thành \x0b
    except Exception:
        t = "".join(r.text for r in p.runs)
    return _lam_sach(t.replace("\x0b", " ").replace("\n", " "))


def _doan_theo_dong(p) -> list[str]:
    """Như _doan_chu nhưng tách tại mỗi chỗ xuống dòng mềm (dùng cho ô tiêu đề)."""
    try:
        t = p.text
    except Exception:
        t = "".join(r.text for r in p.runs)
    return [x for x in (_lam_sach(k) for k in re.split(r"[\x0b\n]", t)) if x]


def tach_tieu_de(text: str) -> tuple[str, str]:
    """Tách một đoạn chữ dài thành (tiêu đề ngắn, phần còn lại cho nội dung).
    Đoạn ngắn thì giữ nguyên làm tiêu đề. Không bao giờ làm mất chữ."""
    text = _lam_sach(text)
    if len(text) <= TIEU_DE_TOI_DA:
        return text, ""
    m = re.match(r"^(.{12,%d}?[.:;!?])\s+(\S.*)$" % TIEU_DE_TOI_DA, text)
    if m:
        return m.group(1).rstrip(".:;").strip(), m.group(2).strip()
    cat = text.rfind(" ", 0, TIEU_DE_TOI_DA - 10)
    cat = cat if cat > 20 else TIEU_DE_TOI_DA - 10
    return text[:cat].rstrip(" ,;:-") + "…", text      # giữ nguyên toàn văn ở nội dung


def _bang(table) -> list[list[str]]:
    """Bảng PowerPoint -> lưới chữ; ô gộp đánh dấu '<' (gộp trái) và '^' (gộp trên)."""
    nr, nc = len(table.rows), len(table.columns)
    luoi = [["" for _ in range(nc)] for _ in range(nr)]
    for i in range(nr):
        for j in range(nc):
            c = table.cell(i, j)
            if c.is_merge_origin:
                luoi[i][j] = _lam_sach(c.text)
                for di in range(c.span_height):
                    for dj in range(c.span_width):
                        if di or dj:
                            if i + di < nr and j + dj < nc:
                                luoi[i + di][j + dj] = "^" if (dj == 0 and di > 0) else "<"
            elif not c.is_spanned:
                luoi[i][j] = _lam_sach(c.text)
    return luoi


def _bang_tu_bieu_do(chart) -> list[list[str]] | None:
    try:
        plot = chart.plots[0]
        cats = [str(c) for c in plot.categories]
        series = list(plot.series)
        if not cats or not series:
            return None
        rows = [[""] + [str(s.name or f"Dãy {k + 1}") for k, s in enumerate(series)]]
        for i, c in enumerate(cats):
            row = [c]
            for s in series:
                vals = list(s.values)
                v = vals[i] if i < len(vals) else ""
                row.append("" if v is None else (f"{v:g}" if isinstance(v, (int, float)) else str(v)))
            rows.append(row)
        return rows
    except Exception:
        return None


def _la_smartart(sh) -> bool:
    try:
        return any("diagram" in (u or "") for u in
                   sh._element.xpath(".//a:graphicData/@uri"))
    except Exception:
        return False


def _duyet(shapes, sl: SlideIn, khoi: list, tieu_de: list, anh: list):
    for sh in shapes:
        try:
            st = sh.shape_type
        except NotImplementedError:
            st = None
        if st == MSO_SHAPE_TYPE.GROUP:
            _duyet(sh.shapes, sl, khoi, tieu_de, anh)
            continue
        if getattr(sh, "has_chart", False) and sh.has_chart:
            rows = _bang_tu_bieu_do(sh.chart)
            if rows:
                sl.tables.append(rows)
                sl.canh_bao.append("có biểu đồ, đã đổi thành bảng số liệu; hãy vẽ lại biểu đồ nếu cần")
            else:
                sl.unsupported.append("biểu đồ")
            continue
        if getattr(sh, "has_table", False) and sh.has_table:
            sl.tables.append(_bang(sh.table))
            continue
        if _la_smartart(sh):
            sl.unsupported.append("SmartArt")
            continue
        if st == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and hasattr(sh, "image")):
            try:
                blob = sh.image.blob
            except Exception:
                sl.canh_bao.append("có hình liên kết tới file ngoài hoặc bị hỏng, chưa chuyển được; "
                                   "hãy chèn lại hình")
                continue
            r = xu_ly(blob)
            if r is None:
                sl.canh_bao.append("có hình định dạng chưa đọc được (thường là EMF/WMF); hãy chèn lại "
                                   "hình dạng PNG/JPG")
                continue
            data, mime, _, _ = r
            anh.append(((sh.top or 0, sh.left or 0), (data, EXT[mime])))
            continue
        if st in (MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
                  MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
            sl.unsupported.append("đối tượng nhúng (video/âm thanh/OLE)")
            continue
        if not sh.has_text_frame:
            continue
        ph = sh.placeholder_format.type if sh.is_placeholder else None
        if ph in _BO_QUA:
            continue
        paras = []
        if ph in _TITLE:
            # ô tiêu đề: mỗi dòng (kể cả xuống dòng mềm) là một mảnh riêng để còn tách tiêu đề khỏi nội dung
            for p in sh.text_frame.paragraphs:
                paras.extend((0, t, False) for t in _doan_theo_dong(p))
        else:
            for p in sh.text_frame.paragraphs:
                t = _doan_chu(p)
                if not t:
                    continue
                gach = bool(_RE_GACH.match(t))
                t = _RE_GACH.sub("", t)
                la_than = sh.is_placeholder and ph not in _TITLE
                paras.append((min(p.level, 1), t, gach or la_than or p.level > 0))
        if not paras:
            continue
        if ph in _TITLE:
            tieu_de.append([t for _, t, _ in paras])
        else:
            khoi.append(((sh.top or 0, sh.left or 0), paras))


def doc_pptx(data: bytes) -> tuple[list[SlideIn], list[str]]:
    prs = Presentation(io.BytesIO(data))
    out: list[SlideIn] = []
    warns: list[str] = []
    for idx, slide in enumerate(prs.slides, 1):
        sl = SlideIn()
        khoi, tieu_de, anh = [], [], []
        _duyet(slide.shapes, sl, khoi, tieu_de, anh)
        khoi.sort(key=lambda x: x[0])
        anh.sort(key=lambda x: x[0])
        sl.images = [x[1] for x in anh]
        phan_them: list = []                              # chữ nằm nhầm trong ô tiêu đề, đưa xuống nội dung
        if tieu_de:
            manh = tieu_de[0]
            sl.title, du = tach_tieu_de(manh[0])
            if du:
                phan_them.append((0, du, False))
            phan_them.extend((0, t, False) for t in manh[1:])
            if phan_them:
                sl.canh_bao.append("ô tiêu đề chứa cả đoạn văn dài hoặc nhiều dòng; đã tách dòng đầu làm "
                                   "tiêu đề, phần còn lại đưa xuống nội dung")
        elif khoi:                                        # không có placeholder tiêu đề
            first = khoi[0][1]
            tt, du = tach_tieu_de(first[0][1])
            if len(first[0][1]) <= TIEU_DE_TOI_DA:
                sl.title = tt
                khoi[0] = (khoi[0][0], first[1:])
            else:
                sl.title = tt
                if du:
                    khoi[0] = (khoi[0][0], [(first[0][0], du, first[0][2])] + first[1:])
        sl.items.extend(phan_them)
        for _, paras in khoi:
            sl.items.extend(paras)
        if slide.has_notes_slide:
            sl.notes = slide.notes_slide.notes_text_frame.text.strip()
        if sl.unsupported:
            warns.append(f"Slide {idx}: có {', '.join(sorted(set(sl.unsupported)))} chưa chuyển "
                         "được, hãy chèn lại thủ công.")
        for w in dict.fromkeys(sl.canh_bao):
            warns.append(f"Slide {idx}: {w}.")
        if not sl.title and (sl.items or sl.images or sl.tables):
            warns.append(f"Slide {idx}: không xác định được tiêu đề, đã để “Slide {idx}”.")
            sl.title = f"Slide {idx}"
        if sl.title or sl.items or sl.images or sl.tables:
            out.append(sl)
    return out, warns
