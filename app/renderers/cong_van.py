# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Dựng công văn theo thể thức Nghị định 30/2020/NĐ-CP (Phụ lục I)."""
from __future__ import annotations

import io
import re
from datetime import date

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

from .. import ban_quyen
from ..hinh import giai_ma
from ..profile import DEFAULT_PROFILE, load_profile
from ..schemas import TEN_LOAI, CongVan
from .common import (ALIGN, add_page_number, bottom_border, cant_split,
                     lap_hang_tieu_de, make_data_table, make_table, set_run_font,
                     text_width_mm, tinh_do_rong_cot, to_nen_o)

QUOC_HIEU = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
TIEU_NGU = "Độc lập - Tự do - Hạnh phúc"


# ------------------------------------------------------ quy tắc chuẩn hóa
def format_ngay(d: date, rules: dict) -> str:
    """Ngày nhỏ hơn 10 và tháng 1, 2 phải ghi thêm số 0 ở trước."""
    day = f"{d.day:02d}" if rules.get("pad_day_below_10") and d.day < 10 else str(d.day)
    month = f"{d.month:02d}" if rules.get("pad_month_1_2") and d.month < 3 else str(d.month)
    return f"ngày {day} tháng {month} năm {d.year}"


def format_so_ky_hieu(so: str, ky_hieu: str, pad: bool) -> str:
    so = so.strip()
    if pad and so.isdigit() and int(so) < 10:
        so = f"{int(so):02d}"
    ky_hieu = ky_hieu.strip().upper()          # ký hiệu viết chữ in hoa
    if not so:
        so = "\u00a0" * 6                      # để trống chỗ điền tay
    return f"Số: {so}/{ky_hieu}" if ky_hieu else f"Số: {so}"


def chuan_hoa_danh_sach(items: list[str]) -> list[str]:
    """Gạch đầu dòng, kết thúc bằng ';' và dòng cuối bằng '.'."""
    sach = []
    for it in items:
        s = re.sub(r"^[\-–—•+*]\s*", "", it.strip()).rstrip(";.,").strip()
        if s:
            sach.append(s)
    last = len(sach) - 1
    return [f"- {s}{'.' if i == last else ';'}" for i, s in enumerate(sach)]


RE_GIUA = re.compile(r"^(QUYẾT ĐỊNH|QUYẾT NGHỊ|NGHỊ QUYẾT|QUY ĐỊNH)\s*:$")
RE_DIEU = re.compile(r"^(Điều\s+\d+[a-zA-Z]?\s*\.)\s*(.*)$")
RE_CAN_CU = re.compile(r"^Căn cứ\b")


RE_HINH = re.compile(r"^\[\[hinh:(\w+)\]\]$")
NGAT_TRANG = "[[ngat-trang]]"          # ngắt trang (giữa các biểu mẫu trong phụ lục)
KHONG_VIEN = "[[khong-vien]]"          # dòng đứng ngay trước một bảng: bảng không viền (khối ký)
RE_HINH_TRONG = re.compile(r"(\[\[hinh:\w+\]\])")
RE_TACH_O = re.compile(r"(?<!\\)\|")
RE_PHAN_CACH = re.compile(r"^:?-{3,}:?$")


def parse_noi_dung(text: str) -> list[tuple[str, str]]:
    """'# ' mục lớn, '## ' mục nhỏ, '- ' gạch đầu dòng; tự nhận 'Căn cứ',
    'Điều N.', 'QUYẾT ĐỊNH:'; '| a | b |' là dòng của bảng, '[[hinh:N]]' là hình;
    còn lại là đoạn văn."""
    blocks: list[tuple[str, str]] = []
    bang: list[str] = []
    cho_khong_vien = [False]

    def xa_bang():
        if bang:
            blocks.append(("bang_kv" if cho_khong_vien[0] else "bang", "\n".join(bang)))
            cho_khong_vien[0] = False
            bang.clear()

    for line in text.splitlines():
        s = line.strip()
        if not s:
            xa_bang()                       # dòng trống ngắt bảng (hai bảng liền nhau là hai bảng)
            continue
        if s == KHONG_VIEN:
            xa_bang()
            cho_khong_vien[0] = True
            continue
        if s.startswith("|"):
            bang.append(s)
            continue
        xa_bang()
        if s == NGAT_TRANG:
            blocks.append(("ngat", ""))
            continue
        m = RE_HINH.match(s)
        if m:
            blocks.append(("hinh", m.group(1)))
        elif s.startswith("## "):
            blocks.append(("tieu_muc", s[3:].strip()))
        elif s.startswith("# "):
            blocks.append(("muc", s[2:].strip()))
        elif s.startswith(("- ", "– ")):
            blocks.append(("gach", "- " + s[2:].strip()))
        elif s.startswith("> "):
            blocks.append(("giua", s[2:].strip()))
        elif RE_GIUA.match(s):
            blocks.append(("giua", s))
        elif RE_DIEU.match(s):
            blocks.append(("dieu", s))
        elif RE_CAN_CU.match(s):
            blocks.append(("can_cu", s))
        else:
            blocks.append(("doan", s))
    xa_bang()
    return blocks


def parse_bang(text: str) -> tuple[list[list[str]], bool]:
    """Dòng '| a | b |' -> (các hàng, có hàng tiêu đề). Dòng '|---|---|' ngay sau hàng đầu = có tiêu đề."""
    rows: list[list[str]] = []
    tieu_de = False
    for ln in text.splitlines():
        s = ln.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|") and not s.endswith("\\|"):
            s = s[:-1]
        cells = [c.strip().replace("\\|", "|") for c in RE_TACH_O.split(s)]
        if len(rows) == 1 and cells and all(RE_PHAN_CACH.match(c) for c in cells):
            tieu_de = True
            continue
        rows.append(cells)
    if not rows:
        return [], False
    n = max(len(r) for r in rows)
    return [r + [""] * (n - len(r)) for r in rows], tieu_de


# ------------------------------------------------------------ bộ dựng
class CongVanBuilder:
    def __init__(self, data: CongVan, prof: dict):
        self.d = data
        self.prof = prof
        self.font = prof["font"]
        self.doc = Document()
        zoom = self.doc.settings.element.find(qn("w:zoom"))           # mẫu mặc định thiếu percent: sửa cho đúng schema
        if zoom is not None and zoom.get(qn("w:percent")) is None:
            zoom.set(qn("w:percent"), "100")
        self._used_cells: set = set()
        self._ngang = False            # đang dựng phần giấy nằm ngang (phụ lục)
        self._trong_phu_luc = False    # đang dựng phụ lục (bảng dùng cỡ chữ nhỏ hơn)
        self._ngat_cho = False         # có ngắt trang chờ áp cho đoạn kế tiếp

    # ---- helpers ----
    def _next_par(self, container):
        if hasattr(container, "_tc"):
            key = container._tc
            first = container.paragraphs[0]
            if key not in self._used_cells and not first.text:
                self._used_cells.add(key)
                return first
        return container.add_paragraph()

    def _p(self, container, text="", *, size, bold=False, italic=False,
           align="left", before=0, after=0, line=1.0, exact=None,
           first=None, left=None, right=None, keep_next=False):
        p = self._next_par(container)
        pf = p.paragraph_format
        p.alignment = ALIGN[align]
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        if exact is not None:
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(exact)
        else:
            pf.line_spacing = line
        if first is not None:
            pf.first_line_indent = first
        if left is not None:
            pf.left_indent = left
        if right is not None:
            pf.right_indent = right
        pf.keep_with_next = keep_next
        pf.widow_control = True
        if self._ngat_cho:               # ngắt trang trước đoạn đầu tiên của khối kế tiếp (kể cả ô đầu của bảng)
            pf.page_break_before = True
            self._ngat_cho = False
        if text:
            self._run(p, text, size, bold, italic)
        return p

    def _run(self, p, text, size, bold=False, italic=False):
        r = p.add_run(text)
        set_run_font(r, self.font, size, bold, italic)
        return r

    def _line(self, container, cell_w_mm, line_w_mm):
        """Đường kẻ ngang, canh giữa, có độ dài line_w_mm."""
        indent = max((cell_w_mm - line_w_mm) / 2, 0)
        p = self._p(container, size=2, exact=3, before=2, left=Mm(indent), right=Mm(indent))
        bottom_border(p)
        return p

    # ---- trang, style ----
    def _setup(self):
        prof = self.prof
        sec = self.doc.sections[0]
        sec.page_width = Mm(prof["page"]["width_mm"])
        sec.page_height = Mm(prof["page"]["height_mm"])
        m = prof["margins_mm"]
        sec.top_margin, sec.bottom_margin = Mm(m["top"]), Mm(m["bottom"])
        sec.left_margin, sec.right_margin = Mm(m["left"]), Mm(m["right"])

        normal = self.doc.styles["Normal"]
        normal.font.name = self.font
        normal.font.size = Pt(prof["body"]["size_pt"])
        rpr = normal.element.get_or_add_rPr()
        rpr.rFonts.set(qn("w:cs"), self.font)
        rpr.rFonts.set(qn("w:eastAsia"), self.font)
        lang = OxmlElement("w:lang")
        lang.set(qn("w:val"), "vi-VN")
        lang.set(qn("w:eastAsia"), "vi-VN")
        rpr.append(lang)

        cp = self.doc.core_properties
        cp.title = f"{TEN_LOAI[self.d.loai_van_ban] or 'Công văn'} {self.d.so}/{self.d.ky_hieu}".strip(" /")
        cp.language = "vi-VN"
        cp.comments = ban_quyen.GHI_CHU

    # ---- khối đầu: cơ quan, quốc hiệu, số, ngày ----
    def _the_thuc_dau(self):
        h, d = self.prof["header"], self.d
        lw, rw = h["left_width_mm"], h["right_width_mm"]
        tbl = make_table(self.doc, 2, [lw, rw])
        for row in tbl.rows:
            cant_split(row)

        # ô 1: tên cơ quan
        left = tbl.cell(0, 0)
        cq = h["co_quan_size_pt"]
        if d.co_quan_chu_quan.strip():
            for ln in d.co_quan_chu_quan.upper().splitlines():
                self._p(left, ln.strip(), size=cq, align="center")
        bh_lines = [x.strip() for x in d.co_quan_ban_hanh.upper().splitlines() if x.strip()]
        for ln in bh_lines:
            self._p(left, ln, size=cq, bold=True, align="center")
        longest = max((text_width_mm(x, cq, True) for x in bh_lines), default=0)
        self._line(left, lw, min(longest, lw) * h["co_quan_line_ratio"])

        # ô 2: quốc hiệu, tiêu ngữ
        right = tbl.cell(0, 1)
        self._p(right, QUOC_HIEU, size=h["quoc_hieu_size_pt"], bold=True, align="center")
        tn = h["tieu_ngu_size_pt"]
        self._p(right, TIEU_NGU, size=tn, bold=True, align="center")
        self._line(right, rw, text_width_mm(TIEU_NGU, tn, True))

        # ô 4: số, ký hiệu + V/v ; ô 3: địa danh, ngày tháng
        so_cell, dd_cell = tbl.cell(1, 0), tbl.cell(1, 1)
        self._p(so_cell, format_so_ky_hieu(d.so, d.ky_hieu, self.prof["so_van_ban"]["pad_below_10"]),
                size=h["so_size_pt"], align="center", before=6)
        if d.loai_van_ban == "cong_van" and d.trich_yeu.strip():
            ty = d.trich_yeu.strip()
            if not re.match(r"(?i)^v/v\b", ty):
                ty = "V/v " + ty
            self._p(so_cell, ty, size=h["trich_yeu_size_pt"], align="center", before=6)

        ngay = format_ngay(d.ngay, self.prof["date_rules"])
        text = f"{d.dia_danh.strip()}, {ngay}" if d.dia_danh.strip() else ngay
        self._p(dd_cell, text, size=h["dia_danh_size_pt"], italic=True,
                align="center", before=6)

    # ---- tên loại + trích yếu (quyết định, tờ trình, thông báo...) ----
    def _ten_loai(self):
        d = self.d
        ten = TEN_LOAI[d.loai_van_ban]
        if ten is None:
            return
        t = self.prof["title_block"]
        self._p(self.doc, ten, size=t["ten_loai_size_pt"], bold=True,
                align="center", before=18, keep_next=True)
        ty = d.trich_yeu.strip()
        if ty:
            sz = t["trich_yeu_size_pt"]
            self._p(self.doc, ty, size=sz, bold=True, align="center",
                    keep_next=True, left=Mm(10), right=Mm(10))
            pg, m = self.prof["page"], self.prof["margins_mm"]
            text_w = pg["width_mm"] - m["left"] - m["right"]
            w = min(text_width_mm(ty, sz, True), text_w - 20)
            self._line(self.doc, text_w, w * t["line_ratio"])

    # ---- kính gửi ----
    def _kinh_gui(self):
        items = [x for x in self.d.kinh_gui if x.strip()]
        if not items:
            return
        k = self.prof["kinh_gui"]
        sz = k["size_pt"]
        if len(items) == 1:
            name = items[0].strip().rstrip(";.,")
            self._p(self.doc, f"Kính gửi: {name}", size=sz, align="center", before=18, after=6)
            return
        hang = Cm(k["hanging_indent_cm"])
        lines = chuan_hoa_danh_sach(items)
        for i, ln in enumerate(lines):
            p = self._p(self.doc, size=sz, before=18 if i == 0 else 0,
                        after=6 if i == len(lines) - 1 else 0,
                        left=hang, first=-hang if i == 0 else None)
            if i == 0:
                p.paragraph_format.tab_stops.add_tab_stop(hang)
                self._run(p, "Kính gửi:\t", sz)
            self._run(p, ln, sz)

    # ---- nội dung ----
    def _noi_dung(self):
        last = self._dung_khoi(self.d.noi_dung)
        if last is not None:            # chữ ký không đứng riêng một trang
            last.paragraph_format.keep_with_next = True

    def _dung_khoi(self, noi_dung: str):
        """Dựng các khối (đoạn, mục, bảng, hình...) từ văn bản đánh dấu; trả về đoạn cuối nếu là đoạn chữ."""
        b = self.prof["body"]
        sz = b["size_pt"]
        last = None
        for kind, text in parse_noi_dung(noi_dung):
            common = dict(size=sz, before=b["space_before_pt"], after=b["space_after_pt"],
                          line=b["line_spacing"])
            if kind == "ngat":
                self._ngat_cho = True
                last = None
            elif kind in ("bang", "bang_kv"):
                self._khoi_bang(text, vien=(kind == "bang"))
                last = None
            elif kind == "hinh":
                self._khoi_hinh(text)
                last = None
            elif kind == "giua":
                last = self._p(self.doc, text, bold=True, align="center", keep_next=True,
                               **{**common, "before": 12, "after": 6})
            elif kind == "dieu":
                m = RE_DIEU.match(text)
                head, rest = m.group(1), m.group(2)
                ca_cau = bool(rest) and (len(text) > 100 or rest.endswith((".", ":", ";", ",")))
                last = self._p(self.doc, align="justify",
                               first=Cm(b["first_line_indent_cm"]), keep_next=not ca_cau, **common)
                if ca_cau:
                    self._run(last, head, sz, bold=True)
                    self._run(last, " " + rest, sz)
                else:
                    self._run(last, text, sz, bold=True)
            else:
                style = {
                    "muc": dict(bold=True, keep_next=True),
                    "tieu_muc": dict(bold=True, italic=True, keep_next=True),
                    "can_cu": dict(italic=True),
                }.get(kind, {})
                last = self._p(self.doc, text, align="justify",
                               first=Cm(b["first_line_indent_cm"]), **common, **style)
        return last

    # ---- bảng và hình trong nội dung ----
    def _be_rong_chu_mm(self) -> float:
        pg, m = self.prof["page"], self.prof["margins_mm"]
        rong = pg["height_mm"] if self._ngang else pg["width_mm"]
        return rong - m["left"] - m["right"]

    def _chen_hinh(self, p, id_, max_w_mm: float, max_h_mm: float) -> bool:
        """Chèn hình vào đoạn p; thu nhỏ cho vừa khung. Trả về False nếu hình hỏng/thiếu."""
        g = giai_ma(self.d.hinh.get(id_) or {}) if self.d.hinh.get(id_) else None
        if g is None:
            return False
        data, w, h = g
        if w > max_w_mm:
            w, h = max_w_mm, h * max_w_mm / w
        if h > max_h_mm:
            w, h = w * max_h_mm / h, max_h_mm
        try:
            p.add_run().add_picture(io.BytesIO(data), width=Mm(w), height=Mm(h))
        except Exception:
            return False
        return True

    def _khoi_hinh(self, id_):
        im = self.prof["image"]
        p = self._p(self.doc, size=self.prof["body"]["size_pt"], align="center",
                    before=im["space_pt"], after=im["space_pt"])
        if not self._chen_hinh(p, id_, self._be_rong_chu_mm(), im["max_height_mm"]):
            p._p.getparent().remove(p._p)     # hình hỏng/thiếu: bỏ dòng tham chiếu

    def _co_chu_bang(self) -> float:
        tb = self.prof["table"]
        return tb.get("phu_luc_size_pt", tb["size_pt"]) if self._trong_phu_luc else tb["size_pt"]

    def _noi_o(self, cell, text, rong_mm, dam):
        tb = self.prof["table"]
        sz = self._co_chu_bang()
        for manh in re.split(r"<br\s*/?>", text):
            for tok in RE_HINH_TRONG.split(manh):
                tok = tok.strip()
                if not tok:
                    continue
                m = RE_HINH.match(tok)
                if m:
                    p = self._p(cell, size=sz, align="center", before=tb["space_pt"], after=tb["space_pt"])
                    if not self._chen_hinh(p, m.group(1), max(rong_mm - 3, 5), self.prof["image"]["max_height_mm"]):
                        p._p.getparent().remove(p._p)
                else:
                    ngan_so = len(tok) <= 4 and tok.replace(".", "").isdigit()
                    self._p(cell, tok, size=sz, bold=dam, align="center" if (dam or ngan_so) else "left",
                            before=tb["space_pt"], after=tb["space_pt"], line=1.0)

    def _khoi_bang(self, text, vien: bool = True):
        rows, tieu_de = parse_bang(text)
        if not rows:
            return
        tb = self.prof["table"]
        nr, nc = len(rows), len(rows[0])
        rong = tinh_do_rong_cot(rows, self._be_rong_chu_mm(), tb["min_col_mm"],
                                co_chu_pt=self._co_chu_bang(), le_mm=2 * tb["cell_pad_mm"] + 1.5)
        t = make_data_table(self.doc, nr, rong, border_eighths=tb["border_eighths_pt"] if vien else 0,
                            cell_pad_mm=tb["cell_pad_mm"])
        for i in range(nr):
            for j in range(nc):
                c = rows[i][j]
                if c in ("<", "^"):
                    continue
                cell = t.cell(i, j)
                self._noi_o(cell, c, rong[j], dam=tieu_de and i == 0)
                if tieu_de and i == 0 and vien:
                    to_nen_o(cell, tb["header_fill"])
        # gộp ô: "<" gộp với ô bên trái, "^" gộp với ô phía trên
        for i in range(nr):
            for j in range(nc):
                if rows[i][j] in ("<", "^"):
                    continue
                w = 1
                while j + w < nc and rows[i][j + w] == "<":
                    w += 1
                h = 1
                while i + h < nr and rows[i + h][j] == "^":
                    h += 1
                if w > 1 or h > 1:
                    t.cell(i, j).merge(t.cell(i + h - 1, j + w - 1))
        if tieu_de:
            lap_hang_tieu_de(t.rows[0])
        # đoạn đệm: tránh Word dính bảng này với bảng liền sau (chữ ký, bảng khác)
        self._p(self.doc, size=4, before=0, after=0, exact=6)

    # ---- chữ ký, nơi nhận ----
    def _chu_ky_noi_nhan(self):
        s, nn, d = self.prof["signature"], self.prof["noi_nhan"], self.d
        sz = s["size_pt"]
        tbl = make_table(self.doc, 1, [s["left_width_mm"], s["right_width_mm"]])
        cant_split(tbl.rows[0])
        left, right = tbl.cell(0, 0), tbl.cell(0, 1)

        items = chuan_hoa_danh_sach(d.noi_nhan)
        if items:
            self._p(left, "Nơi nhận:", size=nn["title_size_pt"], bold=True,
                    italic=True, before=12)
            for it in items:
                self._p(left, it, size=nn["item_size_pt"])

        titles = [t.strip().upper() for t in (d.quyen_han_ky, d.chuc_vu_ky) if t.strip()]
        for i, t in enumerate(titles):
            self._p(right, t, size=sz, bold=True, align="center", before=12 if i == 0 else 0)
        for _ in range(s["blank_lines"]):
            self._p(right, size=sz, align="center")
        if d.nguoi_ky.strip():
            self._p(right, d.nguoi_ky.strip(), size=sz, bold=True, align="center",
                    before=0 if titles else 12)

    # ---- phụ lục (sau chữ ký): trang riêng, có thể khổ ngang ----
    def _phu_luc(self):
        if not self.d.phu_luc.strip():
            return
        prof = self.prof
        pg, m = prof["page"], prof["margins_mm"]
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._ngang = bool(self.d.phu_luc_ngang)
        if self._ngang:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = Mm(pg["height_mm"]), Mm(pg["width_mm"])
        else:
            sec.orientation = WD_ORIENT.PORTRAIT
            sec.page_width, sec.page_height = Mm(pg["width_mm"]), Mm(pg["height_mm"])
        sec.top_margin, sec.bottom_margin = Mm(m["top"]), Mm(m["bottom"])
        sec.left_margin, sec.right_margin = Mm(m["left"]), Mm(m["right"])
        sec.different_first_page_header_footer = False      # phụ lục luôn hiện số trang (đánh số nối tiếp)
        self._trong_phu_luc = True
        self._dung_khoi(self.d.phu_luc)

    def build(self) -> Document:
        self._setup()
        self._the_thuc_dau()
        self._ten_loai()
        self._kinh_gui()
        self._noi_dung()
        self._chu_ky_noi_nhan()
        self._phu_luc()
        pn = self.prof["page_number"]
        add_page_number(self.doc.sections[0], self.font, pn["size_pt"],
                        show_first_page=self.d.hien_so_trang_dau)
        return self.doc


def dung_cong_van(data: CongVan, profile: str = DEFAULT_PROFILE) -> Document:
    return CongVanBuilder(data, load_profile(profile)).build()


def cong_van_to_bytes(data: CongVan, profile: str = DEFAULT_PROFILE) -> bytes:
    buf = io.BytesIO()
    dung_cong_van(data, profile).save(buf)
    return buf.getvalue()
