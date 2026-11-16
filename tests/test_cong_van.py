# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra quy tắc thể thức. Chạy:  python -m pytest -q"""
import io
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx import Document  # noqa: E402

from app.profile import load_profile  # noqa: E402
from app.renderers.cong_van import (chuan_hoa_danh_sach, cong_van_to_bytes,  # noqa: E402
                                    format_ngay, format_so_ky_hieu,
                                    parse_noi_dung)
from app.schemas import CongVan  # noqa: E402

RULES = load_profile("nd30")["date_rules"]


def _mau(**kw):
    base = dict(co_quan_ban_hanh="SỞ A", so="3", ky_hieu="sa-vp", ngay=date(2026, 3, 2),
                dia_danh="Cần Thơ", noi_dung="Nội dung.", noi_nhan=["Như trên", "Lưu: VT"],
                nguoi_ky="Nguyễn Văn A", chuc_vu_ky="GIÁM ĐỐC")
    base.update(kw)
    return CongVan(**base)


def test_ngay_them_so_0():
    assert format_ngay(date(2026, 3, 2), RULES) == "ngày 02 tháng 3 năm 2026"
    assert format_ngay(date(2026, 1, 15), RULES) == "ngày 15 tháng 01 năm 2026"
    assert format_ngay(date(2026, 12, 9), RULES) == "ngày 09 tháng 12 năm 2026"


def test_so_ky_hieu():
    assert format_so_ky_hieu("3", "sgd-vp", True) == "Số: 03/SGD-VP"
    assert format_so_ky_hieu("123", "SGD-VP", True) == "Số: 123/SGD-VP"


def test_danh_sach_noi_nhan():
    assert chuan_hoa_danh_sach(["Như trên;", "- Lưu: VT."]) == ["- Như trên;", "- Lưu: VT."]


def test_parse_noi_dung():
    kinds = [k for k, _ in parse_noi_dung("# A\n## B\n- c\nđoạn")]
    assert kinds == ["muc", "tieu_muc", "gach", "doan"]


def test_trang_giay_le_font():
    doc = Document(io.BytesIO(cong_van_to_bytes(_mau())))
    sec = doc.sections[0]
    assert round(sec.page_width.mm) == 210 and round(sec.page_height.mm) == 297
    assert (round(sec.top_margin.mm), round(sec.bottom_margin.mm),
            round(sec.left_margin.mm), round(sec.right_margin.mm)) == (20, 20, 30, 15)
    assert doc.styles["Normal"].font.name == "Times New Roman"
    assert sec.different_first_page_header_footer is True


def test_co_quoc_hieu_va_tieu_ngu():
    doc = Document(io.BytesIO(cong_van_to_bytes(_mau())))
    text = "\n".join(c.text for t in doc.tables for r in t.rows for c in r.cells)
    assert "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text
    assert "Độc lập - Tự do - Hạnh phúc" in text
    assert "Số: 03/SA-VP" in text
    assert "Cần Thơ, ngày 02 tháng 3 năm 2026" in text
    assert "Nơi nhận:" in text


# ------------------------------------------------ nhập file Word thô
def _docx_tho():
    from docx import Document
    d = Document()
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "UBND TỈNH ABC\nSỞ GIÁO DỤC VÀ ĐÀO TẠO"
    t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"
    t.cell(1, 0).text = "Số: 7/SGDĐT-VP\nV/v hướng dẫn  sinh hoạt ,chuyên môn"
    t.cell(1, 1).text = "Cần Thơ, ngày 5 tháng 1 năm 2026"
    for s in ["Kính gửi: Các trường THPT", "Thực hiện kế hoạch, Sở hướng dẫn như sau:",
              "Nơi nhận:", "- Như trên;", "- Lưu: VT.", "KT. GIÁM ĐỐC", "PHÓ GIÁM ĐỐC",
              "(Đã ký)", "NGUYỄN VĂN A"]:
        d.add_paragraph(s)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def test_nhap_docx_tho():
    from app.importers.docx_reader import doc_docx
    from app.importers.phan_tich import phan_tich
    f, _ = phan_tich(doc_docx(_docx_tho())[0])
    assert f["co_quan_chu_quan"] == "UBND TỈNH ABC"
    assert f["co_quan_ban_hanh"] == "SỞ GIÁO DỤC VÀ ĐÀO TẠO"
    assert (f["so"], f["ky_hieu"]) == ("7", "SGDĐT-VP")
    assert f["ngay"] == "2026-01-05" and f["dia_danh"] == "Cần Thơ"
    assert f["trich_yeu"] == "hướng dẫn sinh hoạt, chuyên môn"
    assert f["kinh_gui"] == ["Các trường THPT"]
    assert f["noi_nhan"] == ["Như trên", "Lưu: VT"]
    assert (f["quyen_han_ky"], f["chuc_vu_ky"], f["nguoi_ky"]) == \
        ("KT. GIÁM ĐỐC", "PHÓ GIÁM ĐỐC", "Nguyễn Văn A")
    assert f["noi_dung"].startswith("Thực hiện kế hoạch")


def test_nhan_dien_quyet_dinh():
    from app.importers.docx_reader import Line
    from app.importers.phan_tich import phan_tich
    ls = [Line("TRƯỜNG THCS LÊ LỢI"), Line("Số: 12/QĐ-THCS"), Line("Lê Lợi, ngày 03 tháng 02 năm 2026"),
          Line("QUYẾT ĐỊNH", True, False, "center"),
          Line("Về việc ban hành Quy chế", True, False, "center"),
          Line("Điều 1. Ban hành quy chế."), Line("Nơi nhận:"), Line("- Lưu: VT."),
          Line("HIỆU TRƯỞNG"), Line("Trần Thị B")]
    f, _ = phan_tich(ls)
    assert f["loai_van_ban"] == "quyet_dinh"
    assert f["trich_yeu"] == "Về việc ban hành Quy chế"
    assert f["co_quan_ban_hanh"] == "TRƯỜNG THCS LÊ LỢI"


def test_so_trang_trang_dau_tuy_chon():
    doc = Document(io.BytesIO(cong_van_to_bytes(_mau(hien_so_trang_dau=True))))
    assert doc.sections[0].different_first_page_header_footer is False


# ------------------------------------------------ slide
def test_chuan_hoa_slide():
    from pptx import Presentation
    from pptx.util import Inches
    from app.importers.pptx_reader import doc_pptx
    from app.renderers.slide_hoc_thuat import ThongTin, dung_slide
    prs = Presentation()
    for ten in ("BÀI 1", "1. Khái niệm", "Tài liệu tham khảo", "2. Ứng dụng", "3. Kết luận"):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)).text_frame
        tb.text = ten
        if ten != "BÀI 1":
            b = s.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4)).text_frame
            b.text = "- Ý một"
            for i in range(10):
                b.add_paragraph().text = f"- Ý số {i}"
    buf = io.BytesIO()
    prs.save(buf)
    slides, _ = doc_pptx(buf.getvalue())
    assert [x.title for x in slides][:2] == ["BÀI 1", "1. Khái niệm"]
    # kiểu học thuật (bài giảng): không có slide kết
    data, warns, n = dung_slide(slides, ThongTin(giang_vien="A", don_vi="Trường B"), profile="slide_hoc_thuat")
    out = Presentation(io.BytesIO(data))
    titles = [sl.shapes.title.text_frame.text for sl in out.slides]
    assert titles[0] == "BÀI 1" and titles[1] == "Nội dung"
    assert titles[-1].startswith("Tài liệu tham khảo")     # đưa xuống cuối
    assert any(t.endswith("(tiếp)") for t in titles)      # tự tách slide dài
    assert n == len(out.slides)
    # kiểu tiêu chuẩn (mặc định): có slide kết, mục lục "Nội dung chính", phông Calibri
    data, warns, n = dung_slide(slides, ThongTin(giang_vien="A", don_vi="Công ty B"))
    out = Presentation(io.BytesIO(data))
    titles = [sl.shapes.title.text_frame.text for sl in out.slides]
    assert titles[0] == "BÀI 1" and titles[1] == "Nội dung chính"
    assert titles[-1] == "Xin trân trọng cảm ơn" and titles[-2].startswith("Tài liệu tham khảo")
    assert any(t.endswith("(tiếp)") for t in titles)
    assert out.slides[2].shapes.title.text_frame.paragraphs[0].runs[0].font.name == "Calibri"
    assert not any("Chưa nhập đủ" in w for w in warns)


# ------------------------------------------------ chữ ký tác giả
def test_chu_ky_tac_gia():
    from app import ban_quyen
    ban_quyen.kiem_tra()                                   # nguyên vẹn thì qua
    doc = Document(io.BytesIO(cong_van_to_bytes(_mau())))
    assert "Thevancorp-1801808400" in doc.core_properties.comments
    old = ban_quyen.CHU_KY
    try:
        ban_quyen.CHU_KY = "Nguoi-khac"
        try:
            ban_quyen.kiem_tra()
            raise AssertionError("sửa chữ ký mà không bị chặn")
        except RuntimeError:
            pass
    finally:
        ban_quyen.CHU_KY = old
