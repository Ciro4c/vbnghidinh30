# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Hàm tiện ích dựng Word bằng python-docx."""
from __future__ import annotations

import os
from functools import lru_cache

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# ---------------------------------------------------------------- font / run
def set_run_font(run, font: str, size_pt: float, bold=False, italic=False):
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    rfonts = run._element.rPr.rFonts
    rfonts.set(qn("w:cs"), font)
    rfonts.set(qn("w:eastAsia"), font)


# ------------------------------------------------- ước lượng bề rộng chữ (mm)
_FONT_PATHS = {
    False: [
        r"C:\Windows\Fonts\times.ttf",
        "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
        "/Library/Fonts/Times New Roman.ttf",
    ],
    True: [
        r"C:\Windows\Fonts\timesbd.ttf",
        "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
        "/Library/Fonts/Times New Roman Bold.ttf",
    ],
}


@lru_cache(maxsize=None)
def _load_font(bold: bool):
    try:
        from PIL import ImageFont
    except ImportError:
        return None
    for path in _FONT_PATHS[bold]:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, 200)
            except OSError:
                continue
    return None


def text_width_mm(text: str, size_pt: float, bold: bool = False) -> float:
    """Bề rộng gần đúng của một dòng chữ. Dùng để kẻ đường gạch dưới
    tiêu ngữ / tên cơ quan có độ dài đúng theo quy định."""
    font = _load_font(bold)
    if font is not None:
        em = font.getlength(text) / 200
    else:  # dự phòng khi không có font: ước lượng thô
        em = 0.0
        for ch in text:
            if ch == " ":
                em += 0.25
            elif ch.isupper():
                em += 0.70
            elif ch.isdigit():
                em += 0.50
            elif ch.isalpha():
                em += 0.47
            else:
                em += 0.30
        if bold:
            em *= 1.05
    return em * size_pt * 25.4 / 72


# ----------------------------------------------------------------- đoạn văn
_PPR_AFTER_PBDR = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)


def bottom_border(paragraph, size_eighths_pt: int = 6):
    """Kẻ đường ngang nét liền dưới đoạn văn (0,75pt)."""
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", str(size_eighths_pt)),
                 ("space", "0"), ("color", "000000")):
        bottom.set(qn("w:" + k), v)
    pbdr.append(bottom)
    pPr.insert_element_before(pbdr, *_PPR_AFTER_PBDR)


# ---------------------------------------------------------------------- bảng
def make_table(doc, rows: int, widths_mm: list[float]):
    """Bảng không viền, độ rộng cố định, không đệm hai bên."""
    t = doc.add_table(rows=rows, cols=len(widths_mm))
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for row in t.rows:
        for cell, w in zip(row.cells, widths_mm):
            cell.width = Mm(w)
    for col, w in zip(t.columns, widths_mm):
        col.width = Mm(w)

    tblPr = t._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(Mm(sum(widths_mm)).twips)))

    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.insert_element_before(mar, "w:tblLook", "w:tblCaption",
                                "w:tblDescription")
    return t


def cant_split(row):
    """Không cho dòng của bảng bị ngắt giữa hai trang."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


# ------------------------------------------------------------- số trang
def _field_run(paragraph, font, size, kind):
    r = paragraph.add_run()
    set_run_font(r, font, size)
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), kind)
    r._r.append(el)


def add_page_number(section, font: str, size_pt: float, show_first_page: bool = False):
    """Số trang canh giữa lề trên. Mặc định từ trang 2 (đúng NĐ 30)."""
    section.different_first_page_header_footer = not show_first_page
    section.header_distance = Mm(10)
    p = section.header.paragraphs[0]
    p.alignment = ALIGN["center"]
    _field_run(p, font, size_pt, "begin")
    r = p.add_run()
    set_run_font(r, font, size_pt)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    r._r.append(instr)
    _field_run(p, font, size_pt, "separate")
    r = p.add_run("2")
    set_run_font(r, font, size_pt)
    _field_run(p, font, size_pt, "end")


# ------------------------------------------------------- bảng dữ liệu có viền
def tinh_do_rong_cot(rows: list[list[str]], tong_mm: float, toi_thieu_mm: float,
                     co_chu_pt: float | None = None, le_mm: float = 0.0) -> list[float]:
    """Chia bề rộng cột theo độ dài nội dung (cột STT hẹp, cột mô tả rộng).
    Mỗi cột luôn đủ rộng cho từ dài nhất trong cột để chữ không bị bẻ giữa từ.
    Cho `co_chu_pt` thì sàn của mỗi cột được đo bằng bề rộng thật của từ dài nhất (in đậm, vì hàng
    tiêu đề đậm) cộng `le_mm` (đệm hai bên ô); nếu tổng sàn vượt bề rộng bảng thì co đều."""
    import re
    nc = max(len(r) for r in rows)
    trong, san = [], []
    for j in range(nc):
        dai, tu, tu_dai = [], 1, ""
        for r in rows:
            c = r[j] if j < len(r) else ""
            if c in ("<", "^"):
                continue
            c = re.sub(r"\[\[hinh:\w+\]\]", "x" * 25, c).replace("<br>", " ")
            dai.append(len(c))
            for w in c.split():
                if len(w) > len(tu_dai):
                    tu_dai = w
            tu = max([tu] + [len(w) for w in c.split()])
        tb = sum(dai) / len(dai) if dai else 4
        co_ban = 0.6 * tb + 0.4 * max(dai or [4])
        trong.append(min(max(co_ban, tu * 1.3 + 2, 4), 60))
        san.append(max(toi_thieu_mm, text_width_mm(tu_dai, co_chu_pt, True) + le_mm)
                   if co_chu_pt else toi_thieu_mm)
    if sum(san) > tong_mm:                       # không đủ chỗ: co đều các sàn
        k = tong_mm / sum(san)
        return [x * k for x in san]
    w = [None] * nc                              # cột nào bị sàn giữ thì cố định, phần còn lại chia theo tỉ lệ
    con_lai = tong_mm
    tu_do = set(range(nc))
    while True:
        tong_t = sum(trong[j] for j in tu_do)
        ap_san = [j for j in tu_do if con_lai * trong[j] / tong_t < san[j]]
        if not ap_san:
            for j in tu_do:
                w[j] = con_lai * trong[j] / tong_t
            break
        for j in ap_san:
            w[j] = san[j]
            con_lai -= san[j]
            tu_do.discard(j)
        if not tu_do:
            break
    k = tong_mm / sum(w)
    return [x * k for x in w]


def make_data_table(doc, n_rows: int, widths_mm: list[float], *, border_eighths: int,
                    cell_pad_mm: float):
    """Bảng có viền đơn, độ rộng cố định, có đệm trong ô."""
    t = doc.add_table(rows=n_rows, cols=len(widths_mm))
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for row in t.rows:
        for cell, w in zip(row.cells, widths_mm):
            cell.width = Mm(w)
        cant_split(row)
    for col, w in zip(t.columns, widths_mm):
        col.width = Mm(w)
    tblPr = t._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(Mm(sum(widths_mm)).twips)))

    bd = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        for k, v in (("val", "single"), ("sz", str(border_eighths)), ("space", "0"), ("color", "000000")):
            el.set(qn("w:" + k), v)
        bd.append(el)
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(Mm(0 if side in ("top", "bottom") else cell_pad_mm).twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    after = ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
    if border_eighths > 0:                       # 0 = bảng không viền (khối ký của biểu mẫu)
        tblPr.insert_element_before(bd, *after)
    tblPr.insert_element_before(mar, "w:tblLook", "w:tblCaption", "w:tblDescription")
    return t


def to_nen_o(cell, mau_hex: str):
    """Tô nền một ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), mau_hex)
    tcPr.append(shd)


def lap_hang_tieu_de(row):
    """Hàng đầu bảng lặp lại ở đầu mỗi trang."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
