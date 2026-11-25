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
    md: str = ""           # cùng chữ với `text` nhưng có đánh dấu **đậm**, *nghiêng* theo từng đoạn chữ (rỗng = không có)


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


_WINGDINGS = {0xF0A8: "☐", 0xF06F: "☐", 0xF0A1: "☐", 0xF071: "❑", 0xF0FE: "☑", 0xF0FD: "☒", 0xF078: "☒", 0xF0FB: "☒",
              0xF0FC: "✓", 0xF0D8: "➢", 0xF0A7: "▪", 0xF0B7: "•", 0xF076: "❖", 0xF0E0: "➔", 0xF0E8: "➔"}


def _ky_tu_sym(el) -> str:
    """Ký tự đặc biệt <w:sym> (Wingdings/Symbol): ô tích, đầu dòng... Không biết thì bỏ qua."""
    try:
        ma = int(el.get(qn("w:char")) or "0", 16)
    except ValueError:
        return ""
    return _WINGDINGS.get(ma if ma >= 0xF000 else ma + 0xF000, "")


def _ky_tu_o_tich(el) -> str:
    """Ô tích dạng biểu mẫu cũ (<w:checkBox>): ☒ nếu đã tích."""
    ck = el.find(qn("w:checked"))
    if ck is not None:
        return "☐" if ck.get(qn("w:val")) in ("0", "false") else "☒"
    mac = el.find(qn("w:default"))
    return "☒" if mac is not None and mac.get(qn("w:val")) not in ("0", "false") else "☐"


def _chu_cua_run(r_el) -> str:
    out = []
    for el in r_el.iterchildren():
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag == qn("w:tab"):
            out.append("\t")
        elif el.tag in (qn("w:br"), qn("w:cr")):
            if el.get(qn("w:type")) != "page":
                out.append("\n")
        elif el.tag == qn("w:sym"):
            out.append(_ky_tu_sym(el))
        elif el.tag == qn("w:noBreakHyphen"):
            out.append("-")
        elif el.tag == qn("w:fldChar"):
            pass
    return "".join(out)


def _text_doan(p_el) -> str:
    out = []
    for el in p_el.iter():
        if _trong(el, p_el, (qn("w:txbxContent"), qn("w:del"), qn("w:instrText"))):
            continue
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag in (qn("w:tab"), qn("w:br"), qn("w:cr")):
            par = el.getparent()
            if par is not None and par.tag == qn("w:r") and not (el.tag == qn("w:br") and el.get(qn("w:type")) == "page"):
                out.append("\t" if el.tag == qn("w:tab") else "\n")
        elif el.tag == qn("w:sym"):
            out.append(_ky_tu_sym(el))
        elif el.tag == qn("w:checkBox"):
            out.append(_ky_tu_o_tich(el) + " ")
    return "".join(out)


def _md_doan(p: Paragraph, text: str, bo_dam: bool = False) -> str:
    """Chữ của đoạn kèm đánh dấu **đậm** / *nghiêng* theo từng đoạn chữ. Trả về "" nếu không có định dạng nào cần giữ
    hoặc không khớp với `text` (khi đó dùng chữ thường)."""
    if "\n" in _text_doan(p._p):
        return ""
    try:
        style_b = p.style.font.bold is True
        style_i = p.style.font.italic is True
    except Exception:
        style_b = style_i = False
    seg: list[list] = []
    for r in p.runs:
        t = _chu_cua_run(r._r).replace("\t", " ")
        if not t:
            continue
        b = r.bold if r.bold is not None else style_b
        i = r.italic if r.italic is not None else style_i
        if bo_dam:
            b = False
        if seg and seg[-1][1] == b and seg[-1][2] == i:
            seg[-1][0] += t
        else:
            seg.append([t, bool(b), bool(i)])
    if not seg or not any(x[1] or x[2] for x in seg) and "*" not in text:
        return ""
    # đoạn gần như cả đoạn cùng một kiểu: bọc cả đoạn một lần (chữ trắng/dấu cách rời không làm vỡ kiểu)
    tong = sum(len(x[0].strip()) for x in seg) or 1
    for k in (1, 2):
        dem = sum(len(x[0].strip()) for x in seg if x[k])
        if dem / tong > 0.8:
            for x in seg:
                x[k] = True
    out = []
    for t, b, i in seg:
        t = t.replace("*", "\\*")
        lead = t[:len(t) - len(t.lstrip())]
        tail = t[len(t.rstrip()):]
        core = t.strip()
        if not core:
            out.append(t)
            continue
        m = "***" if (b and i) else "**" if b else "*" if i else ""
        out.append(f"{lead}{m}{core}{m}{tail}")
    md = lam_sach("".join(out))
    plain = re.sub(r"(?<!\\)\*+", "", md).replace("\\*", "*")
    if re.sub(r"\s+", "", plain) != re.sub(r"\s+", "", text.replace("\t", " ")):
        return ""
    return md


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
            md = _md_doan(p, text) if len(parts) == 1 else ""
            lines.append(Line(text, bold, italic, _ALIGN.get(p.alignment, ""), heading,
                              bullet and k == 0, md=md))
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


def _la_bang_ky_hai_ben(t_el) -> bool:
    """Một hàng có từ hai ô trở lên, mỗi ô mở đầu bằng "ĐẠI DIỆN ...", "BÊN A/B", "NGƯỜI BÁN/MUA..."."""
    for tr in t_el.tr_lst:
        dem = 0
        for tc in tr.tc_lst:
            d = _khong_dau("".join(t.text or "" for t in tc.iter(qn("w:t")))).strip()
            if re.match(r"^(dai dien|ben\s+[a-z]\b|nguoi\s+(ban|mua|giao|nhan|lap|duyet)|chu\s+(dau tu|hop dong))", d):
                dem += 1
        if dem >= 2:
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
    if _la_bang_ky_hai_ben(t_el):
        return False                 # khối ký của hợp đồng/biên bản (ĐẠI DIỆN BÊN A | ĐẠI DIỆN BÊN B): giữ thành bảng
    so_dong = len(t_el.tr_lst)
    so_cot = max((len(tr.tc_lst) for tr in t_el.tr_lst), default=0)
    if not _co_vien(t_el, ten_kieu) and so_cot >= 2 and _la_bang_chu_ky(t_el):
        return False                 # khối ký của biểu mẫu trong phụ lục: giữ nguyên thành bảng không viền
    if not _co_vien(t_el, ten_kieu) and (so_dong <= 2 or so_cot <= 1):
        return True
    return False


def _o_bang(ctx: _Ctx, tc, doc, bo_dam: bool = False) -> str:
    """Nội dung một ô: các đoạn cách nhau bằng <br>, hình thành [[hinh:N]]; giữ **đậm**, *nghiêng* từng đoạn chữ."""
    parts: list[str] = []
    can_le = ""                             # canh lề của ô (đoạn chữ đầu tiên): giữa/phải thì ghi dấu hiệu đầu ô
    trong_cuoi = 0                          # số đoạn trống cuối ô (chỗ để ký tên/đóng dấu)
    for child in tc.iterchildren():
        if child.tag == qn("w:p"):
            par = Paragraph(child, doc)
            if not _text_doan(child).strip() and not any(True for _ in child.iter(_A_BLIP, _V_IMAGEDATA)):
                trong_cuoi += 1
                continue
            trong_cuoi = 0
            if not parts:
                can_le = {WD_ALIGN_PARAGRAPH.CENTER: "[[giua]]", WD_ALIGN_PARAGRAPH.RIGHT: "[[phai]]"}.get(par.alignment, "")
            chu = _text_doan(child)
            nhieu = [lam_sach(x).replace("\t", " ").strip() for x in chu.split("\n")]
            nhieu = [x for x in nhieu if x]
            md = _md_doan(par, nhieu[0], bo_dam) if len(nhieu) == 1 else ""
            if md:
                parts.append(md.replace("\t", " ").strip().replace("|", "\\|"))
            else:
                for x in nhieu:
                    parts.append(x.replace("|", "\\|"))
            parts.extend(f"[[hinh:{i}]]" for i in _hinh_trong_doan(ctx, child))
        elif child.tag == qn("w:tbl"):
            ctx.bao("bang_long", "Có bảng nằm trong ô của bảng khác; nội dung được đọc phẳng, hãy kiểm tra.")
            for t in child.iter(qn("w:p")):
                x = lam_sach(_text_doan(t)).replace("\t", " ").strip()
                if x:
                    parts.append(x.replace("|", "\\|"))
        elif child.tag == qn("w:sdt"):
            for t in child.iter(qn("w:p")):
                x = lam_sach(_text_doan(t)).replace("\t", " ").strip()
                if x:
                    parts.append(x.replace("|", "\\|"))
    if parts and trong_cuoi:
        parts.extend(["[[dong-trong]]"] * min(trong_cuoi, 6))
    return can_le + "<br>".join(parts)


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


def _vi_tri_o(t_el) -> list[list[tuple[float, float]]]:
    """Vị trí ngang (x0, x1, đơn vị twip) của từng ô trong từng hàng. Dựa vào chiều rộng ô (tcW) để các hàng có
    số ô/độ rộng ô khác nhau vẫn thẳng cột như Word hiển thị; thiếu tcW thì dùng lưới (tblGrid + gridSpan)."""
    luoi_cot = []
    try:
        luoi_cot = [int(g.get(qn("w:w")) or 0) for g in t_el.tblGrid.findall(qn("w:gridCol"))]
    except Exception:
        pass
    tong_luoi = sum(luoi_cot)
    hang = []
    for tr in t_el.tr_lst:
        ds, x = [], 0.0
        truoc = tr.grid_before or 0
        x = float(sum(luoi_cot[:truoc])) if luoi_cot else 0.0
        c0 = truoc
        for tc in tr.tc_lst:
            span = tc.grid_span or 1
            w = None
            tcPr = tc.tcPr
            if tcPr is not None and tcPr.find(qn("w:tcW")) is not None:
                e = tcPr.find(qn("w:tcW"))
                if e.get(qn("w:type")) in (None, "dxa") and e.get(qn("w:w")):
                    w = float(e.get(qn("w:w")))
            if w is None or w <= 0:
                w = float(sum(luoi_cot[c0:c0 + span])) if luoi_cot else 1000.0 * span
            ds.append((x, x + w))
            x += w
            c0 += span
        hang.append(ds)
    tong = [h[-1][1] for h in hang if h]
    if tong:
        chuan = max(tong)
        for i, h in enumerate(hang):                # hàng ngắn hơn nhiều so với hàng dài nhất: co giãn cho cùng bề rộng
            if h and abs(h[-1][1] - chuan) > 0.03 * chuan and not (tr_co_truoc(t_el, i)):
                k = chuan / h[-1][1]
                hang[i] = [(a * k, b * k) for a, b in h]
    return hang


def tr_co_truoc(t_el, i) -> bool:
    return bool(t_el.tr_lst[i].grid_before or t_el.tr_lst[i].grid_after)


def _doc_bang(ctx: _Ctx, t_el) -> list[str]:
    vt = _vi_tri_o(t_el)
    # Lưới chuẩn lấy từ hàng "điển hình" (số ô thường gặp nhất, hòa thì hàng nhiều ô hơn): các hàng vẽ lệch nhau
    # (tác giả kéo ô lệch) được ép về cùng các cột đó; ô rộng hơn thì gộp qua nhiều cột chuẩn. Hàng có NHIỀU ô hơn hàng
    # chuẩn (vd. một dòng chia đôi trong bảng toàn ô gộp) thì lưới được chia mịn thêm, tuyệt đối không dồn ô làm mất chữ.
    from collections import Counter
    dem = Counter(len(h) for h in vt if h)
    if not dem:
        return []
    so_o = max(dem, key=lambda k: (dem[k], k))
    tham_chieu = next(h for h in vt if len(h) == so_o)
    bien = [tham_chieu[0][0]] + [b for _, b in tham_chieu]            # các mép cột của hàng chuẩn
    tong_rong = (bien[-1] - bien[0]) or 1.0
    dung_sai = 0.03 * tong_rong
    for ds in vt:
        if len(ds) > so_o:
            for a, b in ds:
                for x in (a, b):
                    if all(abs(x - e) > dung_sai for e in bien):
                        bien.append(x)
    bien.sort()
    tam = [(bien[k] + bien[k + 1]) / 2 for k in range(len(bien) - 1)]
    def chi_so(x):
        return min(range(len(bien)), key=lambda k: abs(bien[k] - x))
    cot_chuan = [(chi_so(a), chi_so(b)) for a, b in tham_chieu]       # cột chuẩn k phủ các cột mịn [i0, i1)
    tieu_de = _hang_tieu_de(t_el, ctx.doc)
    luoi: list[list[str]] = []
    for ri, (tr, ds) in enumerate(zip(t_el.tr_lst, vt)):
        row = [""] * len(tam)
        dau = 0
        for ci, (tc, (a, b)) in enumerate(zip(tr.tc_lst, ds)):
            if len(ds) == so_o:                       # cùng số ô với hàng chuẩn: ô thứ k nằm ở cột chuẩn k
                j0, j1 = cot_chuan[ci]
                j0 = max(j0, dau)
                j1 = max(j1, j0 + 1)
            else:
                trong = [k for k, c in enumerate(tam) if a <= c <= b and k >= dau]
                if trong:
                    j0, j1 = trong[0], trong[-1] + 1
                else:
                    j0 = min(range(dau, len(tam)), key=lambda k: abs(tam[k] - (a + b) / 2)) if dau < len(tam) else len(tam) - 1
                    j1 = j0 + 1
            if j0 >= len(tam):
                break
            j1 = min(j1, len(tam))
            dau = j1
            row[j0] = "^" if tc.vMerge == "continue" else _o_bang(ctx, tc, ctx.doc, bo_dam=(tieu_de and len(t_el.tr_lst) > 1 and ri == 0))
            for j in range(j0 + 1, j1):
                row[j] = "<"
        luoi.append(row)
    n = max((len(r) for r in luoi), default=0)
    if n == 0:
        return []
    out = []
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


# ------------------------------------------------------------ mục lục tự động, ô nội dung (content control)
MUC_LUC = "[[muc-luc]]"


def _la_sdt_muc_luc(sdt) -> bool:
    pr = sdt.find(qn("w:sdtPr"))
    if pr is None:
        return False
    for g in pr.iter(qn("w:docPartGallery")):
        if "table of contents" in (g.get(qn("w:val")) or "").lower():
            return True
    return any("TOC" in (x.text or "").upper().split() or (x.text or "").strip().upper().startswith("TOC")
               for x in sdt.iter(qn("w:instrText")))


def _la_dong_muc_luc(p_el, doc) -> bool:
    """Đoạn thuộc mục lục tự động của Word (kiểu TOC n, hoặc chứa trường TOC)."""
    try:
        ten = (Paragraph(p_el, doc).style.name or "").lower()
    except Exception:
        ten = ""
    if re.match(r"(toc|table of contents)\s*\d", ten) or ten.startswith("mục lục"):
        return True
    return any((x.text or "").strip().upper().startswith("TOC") for x in p_el.iter(qn("w:instrText")))


def _con_cua_than(goc, ctx):
    """Các phần tử con theo thứ tự đọc; content control (w:sdt) được mở ra, mục lục tự động thành dấu [[muc-luc]]."""
    for child in goc.iterchildren():
        if child.tag == qn("w:sdt"):
            ct = child.find(qn("w:sdtContent"))
            if ct is None:
                continue
            if _la_sdt_muc_luc(child):
                ps = list(ct.iterchildren(qn("w:p")))
                # đoạn đầu là tiêu đề của mục lục ("Mục lục", "Nội dung") nếu không phải dòng TOC
                if ps and not _la_dong_muc_luc(ps[0], ctx.doc) and (ps[0].xpath("string(.)") or "").strip():
                    yield ps[0]
                yield _DAU_MUC_LUC
                continue
            yield from _con_cua_than(ct, ctx)
        else:
            yield child


_DAU_MUC_LUC = object()      # phần tử giả báo 'ở đây có mục lục tự động'


# ----------------------------------------------------------------------- vào/ra
def doc_docx_ex(data: bytes) -> tuple[list[Line], list[str], dict]:
    """Đọc file .docx; trả về (các dòng, cảnh báo, hình {id: bản ghi hình})."""
    doc = Document(io.BytesIO(data))
    ctx = _Ctx(doc)
    lines: list[Line] = []
    dau_phan = 0                                  # chỉ số dòng bắt đầu của section hiện tại
    for child in _con_cua_than(doc.element.body, ctx):
        if child is _DAU_MUC_LUC:
            if not lines or lines[-1].raw != MUC_LUC:
                lines.append(Line(DAU_DAC_BIET, kind="muc_luc", raw=MUC_LUC))
            continue
        if child.tag == qn("w:p"):
            if _la_dong_muc_luc(child, doc):
                if not lines or lines[-1].raw != MUC_LUC:
                    lines.append(Line(DAU_DAC_BIET, kind="muc_luc", raw=MUC_LUC))
                continue
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
