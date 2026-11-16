# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra quyết định có phụ lục sau chữ ký, và slide có tiêu đề/đoạn văn dài. Chạy:  python -m pytest -q"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document  # noqa: E402
from docx.enum.section import WD_ORIENT, WD_SECTION  # noqa: E402
from docx.enum.text import WD_BREAK  # noqa: E402
from docx.shared import Mm  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

import _mau  # noqa: E402
from app.importers.docx_reader import doc_docx_ex  # noqa: E402
from app.importers.phan_tich import phan_tich  # noqa: E402
from app.importers.pptx_reader import doc_pptx, tach_tieu_de  # noqa: E402
from app.renderers.cong_van import cong_van_to_bytes, parse_noi_dung  # noqa: E402
from app.renderers.slide_hoc_thuat import ThongTin, dung_slide  # noqa: E402
from app.schemas import CongVan  # noqa: E402


def _quyet_dinh_co_phu_luc() -> bytes:
    d = Document()
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "CÔNG TY ABC\nTỔ KỸ THUẬT"
    t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"
    t.cell(1, 0).text = "Số: 12/QĐ-ABC"
    t.cell(1, 1).text = "Cần Thơ, ngày 5 tháng 1 năm 2026"
    d.add_heading("QUYẾT ĐỊNH", 1)
    d.add_paragraph("Về việc thành lập Tổ Kỹ thuật").runs[0].bold = True
    d.add_paragraph("Căn cứ Luật Doanh nghiệp được Quốc hội nước Cộng hòa xã hội chủ nghĩa Việt Nam "
                    "thông qua ngày 17 tháng 06 năm 2020;")
    d.add_paragraph("QUYẾT ĐỊNH:")
    d.add_heading("Điều 1: Đối tượng áp dụng", 4)
    d.add_paragraph("Thành lập Tổ Kỹ thuật trực thuộc Phòng Hành chính.")
    for s in ["Nơi nhận:", "- Như Điều 1;", "- Lưu: VT.", "GIÁM ĐỐC"]:
        d.add_paragraph(s)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    sec = d.add_section(WD_SECTION.NEW_PAGE)                    # phụ lục khổ ngang
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Mm(297), Mm(210)
    d.add_heading("BẢNG DANH MỤC QUY TRÌNH", 1)
    bt = d.add_table(rows=3, cols=7)
    bt.style = "Table Grid"
    for j in range(7):
        bt.cell(0, j).text = f"Cột {j + 1}"
        bt.cell(1, j).text = "Nội dung dài hơn một chút để thử xuống dòng"
        bt.cell(2, j).text = "x"
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading("Biểu mẫu 02", 3)
    d.add_paragraph("Ghi chú cuối phụ lục.")
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


def _tach_qd():
    lines, _, hinh = doc_docx_ex(_quyet_dinh_co_phu_luc())
    f, w = phan_tich(lines)
    return f, w, hinh


def test_phu_luc_khong_bi_nuot_vao_khoi_chu_ky():
    f, w, _ = _tach_qd()
    assert f["chuc_vu_ky"] == "GIÁM ĐỐC" and f["nguoi_ky"] == ""        # trước đây: cả phụ lục nằm trong chức vụ
    assert f["noi_nhan"] == ["Như Điều 1", "Lưu: VT"]
    assert "BẢNG DANH MỤC QUY TRÌNH" in f["phu_luc"] and "Ghi chú cuối phụ lục" in f["phu_luc"]
    assert f["phu_luc_ngang"] is True
    assert any("phụ lục" in x for x in w)
    khoi = parse_noi_dung(f["phu_luc"])
    assert sum(1 for k, _ in khoi if k == "bang") == 1 and any(k == "ngat" for k, _ in khoi)


def test_can_cu_nhac_toi_ten_nuoc_khong_bi_coi_la_quoc_hieu():
    f, *_ = _tach_qd()
    assert "Căn cứ Luật Doanh nghiệp" in f["noi_dung"]                   # trước đây bị nuốt mất
    assert f["so"] == "12" and f["ngay"] == "2026-01-05"                 # ngày trong câu căn cứ không đè ngày ban hành


def test_dieu_hai_cham_thanh_dieu_cham():
    f, *_ = _tach_qd()
    assert "Điều 1. Đối tượng áp dụng" in f["noi_dung"].splitlines()


def test_dung_phu_luc_o_section_rieng_giay_ngang():
    f, _, hinh = _tach_qd()
    f["hinh"] = hinh
    doc = Document(io.BytesIO(cong_van_to_bytes(CongVan(**f))))
    assert len(doc.sections) == 2
    assert doc.sections[0].orientation == WD_ORIENT.PORTRAIT
    assert doc.sections[1].orientation == WD_ORIENT.LANDSCAPE
    assert doc.sections[1].page_width > doc.sections[1].page_height
    ky = doc.tables[-2]                                                 # bảng chữ ký (bảng cuối của phần chính)
    assert len(ky.cell(0, 1).text) < 40                                 # khối ký chỉ có chức vụ, không có phụ lục
    bang_pl = doc.tables[-1]
    assert len(bang_pl.columns) == 7
    rong = sum(c.width for c in bang_pl.columns) / 36000
    assert rong > 200                                                   # bảng dùng hết bề rộng giấy ngang (252 mm)


def test_khong_co_phu_luc_thi_van_mot_section():
    d = _mau.docx_co_bang_hinh()
    f, _ = phan_tich(doc_docx_ex(d)[0])
    assert f["phu_luc"] == "" and f["phu_luc_ngang"] is False
    f["hinh"] = doc_docx_ex(d)[2]
    assert len(Document(io.BytesIO(cong_van_to_bytes(CongVan(**f)))).sections) == 1


# ------------------------------------------------------------------ slide
DOAN_DAI = ("Đầu phun sprinkler thông thường có chốt chặn (cap) được giữ cố định bằng một cơ chế kích hoạt. "
            "Loại cơ chế phổ biến nhất là ống thủy tinh chứa chất lỏng giãn nở khi gặp nhiệt độ nóng. "
            "Chất lỏng này được thiết kế để giãn nở ở một nhiệt độ nhất định và làm vỡ ống thủy tinh. "
            "Khi đó dưới áp lực nước phía sau làm cho chốt giữ rơi ra ngoài, cho phép nước phun ra.")


def _pptx_tieu_de_chua_doan_van() -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text_frame.text = "Nguyên lý hoạt động của đầu phun Sprinkler\v" + DOAN_DAI   # xuống dòng mềm
    s.shapes.add_picture(_mau.png(900, 300), Inches(1), Inches(3.5), Inches(8), Inches(2.5))
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text_frame.text = "HỆ THỐNG PCCC"
    tb = s2.shapes.add_textbox(Inches(1), Inches(1.6), Inches(8), Inches(3))
    tb.text_frame.text = (DOAN_DAI + " ") * 6
    b = io.BytesIO()
    prs.save(b)
    return b.getvalue()


def test_xuong_dong_mem_khong_dinh_chu_va_tieu_de_dai_duoc_tach():
    slides, w = doc_pptx(_pptx_tieu_de_chua_doan_van())
    s = slides[0]
    assert s.title == "Nguyên lý hoạt động của đầu phun Sprinkler"      # trước đây: "...SprinklerĐầu phun..."
    assert s.items and s.items[0][1].startswith("Đầu phun sprinkler") and len(s.images) == 1
    assert any("ô tiêu đề" in x for x in w)


def test_tach_tieu_de_khong_mat_chu():
    assert tach_tieu_de("Ngắn gọn") == ("Ngắn gọn", "")
    tt, du = tach_tieu_de("Nguyên lý hoạt động của hệ thống. " + "chữ " * 60)
    assert tt == "Nguyên lý hoạt động của hệ thống" and du.startswith("chữ")
    tt, du = tach_tieu_de("từ " * 80)                                    # không có dấu câu: cắt ở ranh giới từ, giữ toàn văn
    assert tt.endswith("…") and len(tt) <= 100 and du.startswith("từ từ")


def test_slide_tieu_de_gon_va_chu_khong_tran():
    slides, _ = doc_pptx(_pptx_tieu_de_chua_doan_van())
    out, _, n = dung_slide(slides, ThongTin("BÀI GIẢNG", "", "", "", ""))
    prs = Presentation(io.BytesIO(out))
    cao = prs.slide_height
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_text_frame and sh == sl.shapes.title and sh.text_frame.text:
                assert len(sh.text_frame.text) <= 110                     # tiêu đề không còn là cả đoạn văn
            assert sh.top + sh.height <= cao + 1
    # đoạn 1.000+ ký tự được tách sang nhiều slide thay vì tràn khỏi slide
    sl_pccc = [s for s in prs.slides if s.shapes.title and s.shapes.title.text_frame.text.startswith("HỆ THỐNG PCCC")]
    assert len(sl_pccc) >= 3
