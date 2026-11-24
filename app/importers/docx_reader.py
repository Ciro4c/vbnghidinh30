# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Đọc file Word/TXT thô thành danh sách dòng kèm gợi ý định dạng.

Ngoài chữ, còn giữ lại:
- Bảng dữ liệu (có viền): thành các dòng "| ô | ô |", ô gộp ghi "<" (gộp với ô bên trái)
  và "^" (gộp với ô phía trên). Bảng dùng để dàn trang (đầu văn bản, khối chữ ký) vẫn được
  đọc phẳng như trước.
- Hình ảnh: thành dòng "[[hinh:N]]", dữ liệu hình trả về riêng (đã thu nhỏ nếu quá lớn).
- Hộp văn bản (text box): đọc nội dung ngay sau đoạn chứa hộp.
"""
from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from ..hinh import KhoHinh

_ALIGN = {
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    WD_ALIGN_PARAGRAPH.LEFT: "left",
}

_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
_V_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"

DAU_DAC_BIET = "§"      # chữ thay thế cho dòng bảng/hình để các bước nhận diện không bị nhiễu


@dataclass
class Line:
    text: str
    bold: bool = False
    italic: bool = False
    align: str = ""
    heading: int = 0       # cấp Heading nếu đoạn dùng style Heading
    bullet: bool = False   # đoạn có đánh số/gạch đầu dòng tự động của Word
    kind: str = "text"     # "text" | "bang" | "hinh" | "ngat" (ngắt trang)
    raw: str = ""          # nội dung đánh dấu thật của dòng bảng/hình/ngắt trang
    ngang: bool = False    # dòng nằm trong phần (section) giấy nằm ngang


_LATEX = {r"\phi": "φ", r"\varphi": "φ", r"\le": "≤", r"\leq": "≤", r"\ge": "≥", r"\geq": "≥",
          r"\pm": "±", r"\times": "×", r"\approx": "≈", r"\neq": "≠", r"\Delta": "Δ", r"\Omega": "Ω",
          r"\mu": "μ", r"\circ": "°", r"\rightarrow": "→", r"\%": "%", r"\cdot": "·", r"\sim": "~"}
_RE_LATEX = re.compile(r"\$([^$\n]{1,60})\$")


def _latex_don_gian(s: str) -> str:
    """Công thức LaTeX đơn giản dán từ chat AI ($\\phi$, $\\le 15$, $^\\circ\\text{C}$) -> ký tự thật."""
    def doi(m):
        t = m.group(1)
        if "\\" not in t and "^" not in t:          # \"$5 và $7\": không phải công thức
            return m.group(0)
        for k in sorted(_LATEX, key=len, reverse=True):
            t = re.sub(re.escape(k) + r"(?![A-Za-z])", _LATEX[k], t)
        t = re.sub(r"\\text\{([^}]*)\}", r"\1", t)
        t = t.replace("^2", "²").replace("^3", "³").replace("^°", "°").replace("{", "").replace("}", "")
        return re.sub(r"\s+", " ", t).strip()
    return _RE_LATEX.sub(doi, s)


def lam_sach(s: str) -> str:
    """Chuẩn hóa khoảng trắng và dấu câu."""
    s = s.replace("\u00a0", " ").replace("\u200b", "")
    if "$" in s:
        s = _latex_don_gian(s)
    s = re.sub(r" {3,}|\t+", "\t", s)          # khoảng cách lớn = chia cột trái/phải
    s = re.sub(r" {2,}", " ", s).strip()
    s = re.sub(r"[ ]+([,;:.!?)])", r"\1", s)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r",(?=[^\s\d])", ", ", s)
    return s


def _khong_dau(s: str) -> str:
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


class _Ctx:
    """Trạng thái dùng chung khi đọc một file: kho hình, danh sách cảnh báo."""

    def __init__(self, doc):
        self.doc = doc
        self.kho = KhoHinh()
        self.warns: list[str] = []
        self._da_bao: set[str] = set()
        self.so_bang = 0

    def bao(self, khoa: str, msg: str):
        if khoa not in self._da_bao:
            self._da_bao.add(khoa)
            self.warns.append(msg)


def _trong(el, goc, tags) -> bool:
    """el có nằm trong phần tử nào đó thuộc `tags`, tính từ el lên tới `goc` không."""
    for a in el.iterancestors():
        if a is goc:
            return False
        if a.tag in tags:
            return True
    return False


def _text_doan(p_el) -> str:
    out = []
    for el in p_el.iter():
        if _trong(el, p_el, (qn("w:txbxContent"),)):
            continue
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag in (qn("w:tab"), qn("w:br"), qn("w:cr")):
            par = el.getparent()
            if par is not None and par.tag == qn("w:r"):
                out.append("\t" if el.tag == qn("w:tab") else "\n")
    return "".join(out)


def _ty_le(p: Paragraph, thuoc_tinh: str) -> bool:
    tong = dem = 0
    try:
        style_val = getattr(p.style.font, thuoc_tinh) is True
    except Exception:
        style_val = False
    for r in p.runs:
        n = len(r.text.strip())
        if not n:
            continue
        v = getattr(r, thuoc_tinh)
        tong += n
        if v is True or (v is None and style_val):
            dem += n
    return tong > 0 and dem / tong > 0.8


# --------------------------------------------------------------------- hình ảnh
def _hinh_trong_doan(ctx: _Ctx, p_el) -> list[str]:
    """Lấy các hình nằm trực tiếp trong đoạn (không tính hình trong hộp văn bản, bản dự phòng VML)."""
    ids = []
    for el in p_el.iter(_A_BLIP, _V_IMAGEDATA):
        if _trong(el, p_el, (_MC_FALLBACK, qn("w:txbxContent"))):
            continue
        rid = el.get(_R + "embed") or el.get(_R + "id")
        if not rid:
            if el.get(_R + "link"):
                ctx.bao("hinh_lien_ket", "Có hình chèn dạng liên kết tới file ngoài; không chuyển được, "
                                         "hãy chèn lại hình vào văn bản.")
            continue
        try:
            part = ctx.doc.part.related_parts[rid]
        except KeyError:
            continue
        w_mm = 0.0
        for a in el.iterancestors():
            if a.tag == qn("w:drawing"):
                ext = a.find(".//" + _WP_EXTENT)
                if ext is not None and ext.get("cx", "").isdigit():
                    w_mm = int(ext.get("cx")) / 36000
                break
        id_ = ctx.kho.them(part.blob, w_mm)
        if id_:
            ids.append(id_)
        else:
            ctx.bao("hinh_loi", "Có hình định dạng chưa đọc được (thường là EMF/WMF hoặc hình hỏng); "
                                "hãy chèn lại hình dạng PNG/JPG.")
    return ids


def _hop_van_ban(p_el):
    """Các hộp văn bản ngoài cùng của đoạn, bỏ bản trùng lặp ở nhánh dự phòng của Word."""
    return [tb for tb in p_el.iter(qn("w:txbxContent"))
            if not _trong(tb, p_el, (_MC_FALLBACK, qn("w:txbxContent")))]


# ------------------------------------------------------------------------- đoạn
NGAT_TRANG = "[[ngat-trang]]"
KHONG_VIEN = "[[khong-vien]]"


def _co_ngat_trang(p_el) -> tuple[bool, bool]:
    """(có ngắt trang, ngắt trang đứng trước chữ của đoạn)."""
    pPr = p_el.pPr
    truoc = pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None
    co = truoc
    for el in p_el.iter(qn("w:br"), qn("w:t")):
        if _trong(el, p_el, (qn("w:txbxContent"),)):
            continue
        if el.tag == qn("w:br") and el.get(qn("w:type")) == "page":
            if not co:
                truoc = True
            co = True
            break
        if el.tag == qn("w:t") and (el.text or "").strip():
            break
    return co, truoc


def _them_doan(lines: list[Line], p: Paragraph, ctx: _Ctx):
    co_ngat, ngat_truoc = _co_ngat_trang(p._p)
    if co_ngat and ngat_truoc:
        lines.append(Line(DAU_DAC_BIET, kind="ngat", raw=NGAT_TRANG))
    parts = [lam_sach(x) for x in _text_doan(p._p).split("\n")]
    parts = [x for x in parts if x]
    if parts:
        heading = 0
        try:
            name = (p.style.name or "").lower()
            m = re.match(r"heading\s*(\d)", name)
            if m:
                heading = int(m.group(1))
        except Exception:
            pass
        pPr = p._p.pPr
        bullet = pPr is not None and pPr.numPr is not None
        bold, italic = _ty_le(p, "bold"), _ty_le(p, "italic")
        for k, text in enumerate(parts):
            lines.append(Line(text, bold, italic, _ALIGN.get(p.alignment, ""), heading,
                              bullet and k == 0))
    for tb in _hop_van_ban(p._p):
        ctx.bao("hop_van_ban", "File có hộp văn bản (text box); nội dung trong hộp được đọc ngay sau "
                               "đoạn chứa hộp, hãy kiểm tra thứ tự.")
        for child in tb.iterchildren(qn("w:p")):
            _them_doan(lines, Paragraph(child, p._parent), ctx)
    for id_ in _hinh_trong_doan(ctx, p._p):
        lines.append(Line(DAU_DAC_BIET, kind="hinh", raw=f"[[hinh:{id_}]]"))
    if co_ngat and not ngat_truoc:
        lines.append(Line(DAU_DAC_BIET, kind="ngat", raw=NGAT_TRANG))


# ------------------------------------------------------------------------- bảng
_TEN_KIEU_THUONG = ("normal table", "table normal", "")


def _co_vien(t_el, ten_kieu: str) -> bool:
    def that(el):
        return el is not None and any(c.get(qn("w:val")) not in (None, "nil", "none") for c in el)

    tblPr = t_el.tblPr
    tb = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
    if that(tb):
        return True
    if any(that(x) for x in t_el.iter(qn("w:tcBorders"))):
        return True
    if tb is not None:                       # có khai báo viền nhưng toàn "none"
        return False
    return ten_kieu not in _TEN_KIEU_THUONG


def _la_bang_chu_ky(t_el) -> bool:
    """Bảng có ô dạng \"(Ký và ghi rõ họ tên)\": khối ký của biểu mẫu, không phải khối ký của văn bản."""
    for tc in t_el.iter(qn("w:tc")):
        d = _khong_dau("".join(t.text or "" for t in tc.iter(qn("w:t")))).strip()
        if re.search(r"\(\s*ky\b", d):
            return True
    return False


def _la_bang_bo_cuc(t_el, ten_kieu: str) -> bool:
    """Bảng dàn trang (đầu văn bản, khối chữ ký...) thì đọc phẳng; bảng dữ liệu thì giữ cấu trúc."""
    nt = _khong_dau(" ".join(t.text or "" for t in t_el.iter(qn("w:t"))))
    if "cong hoa xa hoi" in nt:
        return True
    for tc in t_el.iter(qn("w:tc")):
        d = _khong_dau("".join(t.text or "" for t in tc.iter(qn("w:t")))).strip()
        if re.match(r"^noi nhan\s*:?", d) or re.match(r"^(tm|kt|tl|tuq|q)\.", d):
            return True
    so_dong = len(t_el.tr_lst)
    so_cot = max((len(tr.tc_lst) for tr in t_el.tr_lst), default=0)
    if not _co_vien(t_el, ten_kieu) and so_cot >= 2 and _la_bang_chu_ky(t_el):
        return False                 # khối ký của biểu mẫu trong phụ lục: giữ nguyên thành bảng không viền
    if not _co_vien(t_el, ten_kieu) and (so_dong <= 2 or so_cot <= 1):
        return True
    return False


def _o_bang(ctx: _Ctx, tc, doc) -> str:
    """Nội dung một ô: các đoạn cách nhau bằng <br>, hình thành [[hinh:N]]."""
    parts: list[str] = []
    for child in tc.iterchildren():
        if child.tag == qn("w:p"):
            for x in _text_doan(child).split("\n"):
                x = lam_sach(x).replace("\t", " ").strip()
                if x:
                    parts.append(x.replace("|", "\\|"))
            parts.extend(f"[[hinh:{i}]]" for i in _hinh_trong_doan(ctx, child))
        elif child.tag == qn("w:tbl"):
            ctx.bao("bang_long", "Có bảng nằm trong ô của bảng khác; nội dung được đọc phẳng, hãy kiểm tra.")
            for t in child.iter(qn("w:p")):
                x = lam_sach(_text_doan(t)).replace("\t", " ").strip()
                if x:
                    parts.append(x.replace("|", "\\|"))
    return "<br>".join(parts)


def _hang_tieu_de(t_el, doc) -> bool:
    if not t_el.tr_lst:
        return False
    tr = t_el.tr_lst[0]
    trPr = tr.trPr
    if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
        return True
    tong = dam = 0
    for p in tr.iter(qn("w:p")):
        par = Paragraph(p, doc)
        n = len(par.text.strip())
        if n:
            tong += 1
            dam += 1 if _ty_le(par, "bold") else 0
    return tong > 0 and dam / tong >= 0.8


def _doc_bang(ctx: _Ctx, t_el) -> list[str]:
    luoi: list[list[str]] = []
    for tr in t_el.tr_lst:
        row = [""] * (tr.grid_before or 0)
        for tc in tr.tc_lst:
            span = tc.grid_span or 1
            row.append("^" if tc.vMerge == "continue" else _o_bang(ctx, tc, ctx.doc))
            row.extend(["<"] * (span - 1))
        row.extend([""] * (tr.grid_after or 0))
        luoi.append(row)
    n = max((len(r) for r in luoi), default=0)
    if n == 0:
        return []
    out = []
    tieu_de = _hang_tieu_de(t_el, ctx.doc)
    for i, row in enumerate(luoi):
        row = row + [""] * (n - len(row))
        out.append("| " + " | ".join(row) + " |")
        if i == 0 and tieu_de and len(luoi) > 1:
            out.append("|" + "|".join(["---"] * n) + "|")
    ctx.so_bang += 1
    return out


def _them_bang(lines: list[Line], t_el, ctx: _Ctx):
    try:
        ten_kieu = (Table(t_el, ctx.doc).style.name or "").lower()
    except Exception:
        ten_kieu = ""
    if _la_bang_bo_cuc(t_el, ten_kieu):
        t = Table(t_el, ctx.doc)
        for row in t.rows:
            seen = set()
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                for p in cell.paragraphs:
                    _them_doan(lines, p, ctx)
        return
    dong = _doc_bang(ctx, t_el)
    if dong:
        if not _co_vien(t_el, ten_kieu):
            dong.insert(0, KHONG_VIEN)        # bảng không viền (khối ký của biểu mẫu)
        lines.append(Line(DAU_DAC_BIET, kind="bang", raw="\n".join(dong) + "\n"))


def _danh_dau_ngang(lines: list[Line], tu: int, sect) -> None:
    """Section này có giấy nằm ngang không (khai báo orient hoặc rộng > cao)."""
    sz = sect.find(qn("w:pgSz"))
    if sz is None:
        return
    try:
        ngang = sz.get(qn("w:orient")) == "landscape" or int(sz.get(qn("w:w"), 0)) > int(sz.get(qn("w:h"), 0))
    except ValueError:
        ngang = False
    if ngang:
        for ln in lines[tu:]:
            ln.ngang = True


# ----------------------------------------------------------------------- vào/ra
def doc_docx_ex(data: bytes) -> tuple[list[Line], list[str], dict]:
    """Đọc file .docx; trả về (các dòng, cảnh báo, hình {id: bản ghi hình})."""
    doc = Document(io.BytesIO(data))
    ctx = _Ctx(doc)
    lines: list[Line] = []
    dau_phan = 0                                  # chỉ số dòng bắt đầu của section hiện tại
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            _them_doan(lines, Paragraph(child, doc), ctx)
            pPr = child.pPr
            sect = pPr.find(qn("w:sectPr")) if pPr is not None else None
            if sect is not None:                  # hết một section: đánh dấu giấy ngang cho các dòng của nó
                _danh_dau_ngang(lines, dau_phan, sect)
                dau_phan = len(lines)
        elif child.tag == qn("w:tbl"):
            _them_bang(lines, child, ctx)
        elif child.tag == qn("w:sectPr"):
            _danh_dau_ngang(lines, dau_phan, child)
    xml = doc.element.xml
    if "<w:ins " in xml or "<w:del " in xml:
        ctx.warns.append("File còn các thay đổi đang theo dõi (track changes); chỉ đọc phần văn bản "
                         "hiện hành. Nên chấp nhận/loại bỏ thay đổi trước khi chuyển.")
    return lines, ctx.warns, ctx.kho.hinh


def doc_docx(data: bytes) -> tuple[list[Line], list[str]]:
    lines, warns, _ = doc_docx_ex(data)
    return lines, warns


def doc_txt(data: bytes) -> tuple[list[Line], list[str]]:
    for enc in ("utf-8-sig", "utf-16", "cp1258"):
        try:
            text = data.decode(enc)
            break
        except UnicodeError:
            continue
    else:
        text = data.decode("latin-1")
    lines = [Line(lam_sach(x)) for x in text.splitlines() if lam_sach(x)]
    return lines, []
