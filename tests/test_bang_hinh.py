# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra bảng/hình trong Word và PowerPoint. Chạy:  python -m pytest -q"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document  # noqa: E402

import _mau  # noqa: E402
from app.importers.docx_reader import doc_docx, doc_docx_ex  # noqa: E402
from app.importers.phan_tich import phan_tich  # noqa: E402
from app.renderers.cong_van import cong_van_to_bytes, parse_bang, parse_noi_dung  # noqa: E402
from app.schemas import CongVan  # noqa: E402


def _tach():
    lines, warns, hinh = doc_docx_ex(_mau.docx_co_bang_hinh())
    f, w = phan_tich(lines)
    return lines, warns, hinh, f, w


def test_bang_bo_cuc_van_doc_phang_nhu_cu():
    _, _, _, f, _ = _tach()
    assert f["co_quan_ban_hanh"] == "SỞ GIÁO DỤC VÀ ĐÀO TẠO"
    assert (f["so"], f["ky_hieu"]) == ("7", "SGDĐT-VP")
    assert f["nguoi_ky"] == "Nguyễn Văn A" and f["noi_nhan"] == ["Như trên", "Lưu: VT"]
    assert "CỘNG HÒA" not in f["noi_dung"] and "Số:" not in f["noi_dung"]


def test_bang_du_lieu_giu_cau_truc_va_o_gop():
    *_, f, _ = _tach()
    khoi = parse_noi_dung(f["noi_dung"])
    bang = [t for k, t in khoi if k == "bang"]
    assert len(bang) == 1
    rows, tieu_de = parse_bang(bang[0])
    assert tieu_de and rows[0] == ["STT", "Nội dung công việc", "Đơn vị", "Ghi chú"]
    assert rows[1][3] == "Gộp dọc" and rows[2][3] == "^"
    assert rows[4][1] == "Gộp ngang" and rows[4][2] == "<"
    assert rows[1][2] == "Phòng A<br>Phòng B"


def test_hinh_va_hop_van_ban():
    _, warns, hinh, f, _ = _tach()
    assert len(hinh) == 1 and "[[hinh:1]]" in f["noi_dung"]
    assert "Chữ trong hộp" in f["noi_dung"]
    assert any("hộp văn bản" in w for w in warns)


def test_dung_lai_word_co_bang_hinh():
    _, _, hinh, f, _ = _tach()
    f["hinh"] = hinh
    doc = Document(io.BytesIO(cong_van_to_bytes(CongVan(**f))))
    co_vien = [t for t in doc.tables if t._tbl.tblPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders") is not None]
    assert len(co_vien) == 1
    t = co_vien[0]
    assert len(t.rows) == 5 and len(t.columns) == 4
    assert t.cell(1, 3)._tc is t.cell(2, 3)._tc          # gộp dọc
    assert t.cell(4, 1)._tc is t.cell(4, 2)._tc          # gộp ngang
    assert len(doc.inline_shapes) == 1
    assert doc.inline_shapes[0].width.mm <= 165.5


def test_hinh_thieu_thi_bo_qua_khong_loi():
    *_, f, _ = _tach()
    f["hinh"] = {}
    Document(io.BytesIO(cong_van_to_bytes(CongVan(**f))))   # không ném lỗi


def test_doc_docx_van_tra_hai_gia_tri():
    lines, warns = doc_docx(_mau.docx_co_bang_hinh())
    assert lines and isinstance(warns, list)


def test_hinh_lon_duoc_thu_nho():
    from app.hinh import MAX_CANH, KhoHinh
    kho = KhoHinh()
    i = kho.them(_mau.png(5000, 3000).getvalue(), 100)
    assert i and max(kho.hinh[i]["w_px"], kho.hinh[i]["h_px"]) <= MAX_CANH
    assert kho.them(b"khong-phai-hinh") is None and kho.loi == 1


# ------------------------------------------------ PowerPoint
def test_slide_bang_dai_tach_trang_va_gop_o():
    from pptx import Presentation
    from pptx.util import Inches
    from app.importers.pptx_reader import doc_pptx
    from app.renderers.slide_hoc_thuat import ThongTin, dung_slide
    slides, _ = doc_pptx(_mau.pptx_co_bang_hinh())
    bang = slides[1].tables[0]
    assert bang[0][1] == "<"                               # ô gộp được giữ
    data, warns, n = dung_slide(slides, ThongTin(giang_vien="A", don_vi="B"))
    out = Presentation(io.BytesIO(data))
    tieu_de = [s.shapes.title.text_frame.text for s in out.slides]
    assert sum(t.startswith("Bảng dài") for t in tieu_de) >= 2      # tách nhiều slide
    cao = out.slide_height
    for s in out.slides:
        for sh in s.shapes:
            if sh.has_table:
                assert sh.top + sh.height <= cao, "bảng tràn khỏi slide"
                assert sh.table.cell(0, 0).text.startswith("R0C0")   # lặp lại hàng tiêu đề
    assert any("tách thành" in w for w in warns)


def test_slide_nhieu_hinh_xep_luoi_va_thu_nho():
    from pptx import Presentation
    from app.importers.pptx_reader import doc_pptx
    from app.renderers.slide_hoc_thuat import ThongTin, dung_slide
    slides, _ = doc_pptx(_mau.pptx_co_bang_hinh())
    assert len(slides[2].images) == 5
    data, _, _ = dung_slide(slides, ThongTin(giang_vien="A", don_vi="B"))
    out = Presentation(io.BytesIO(data))
    sl = [s for s in out.slides if s.shapes.title.text_frame.text.startswith("Năm ảnh")]
    assert len(sl) == 2                                       # 4 hình + 1 hình
    assert len(data) < 1_500_000                              # hình 4000x3000 đã thu nhỏ


def test_slide_hinh_hong_chi_canh_bao():
    import app.importers.pptx_reader as R
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Hình hỏng"
    s.shapes.add_picture(_mau.png(), Inches(1), Inches(2), Inches(2))
    b = io.BytesIO()
    prs.save(b)
    cu = R.xu_ly
    R.xu_ly = lambda blob: None
    try:
        slides, warns = R.doc_pptx(b.getvalue())
    finally:
        R.xu_ly = cu
    assert slides and any("chưa đọc được" in w for w in warns)
