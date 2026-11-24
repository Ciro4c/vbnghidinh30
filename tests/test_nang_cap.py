# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra các tính năng nâng cấp: thành phần thể thức, ký tự lạ, chữ ký/mộc, PDF, slide, lớp API."""
import base64
import io
import json
import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx import Document  # noqa: E402
from PIL import Image, ImageDraw  # noqa: E402

from app import api_core, khung, ky_tu_la, slide_lint  # noqa: E402
from app.importers.pptx_reader import SlideIn  # noqa: E402
from app.profile import ap_dung_tuy_chinh, load_profile  # noqa: E402
from app.renderers.cong_van import cong_van_to_bytes_ex  # noqa: E402
from app.renderers.slide_hoc_thuat import ThongTin, dung_slide  # noqa: E402
from app.schemas import CongVan  # noqa: E402


def _cv(**kw):
    base = dict(co_quan_ban_hanh="SỞ A", so="3", ky_hieu="sa-vp", ngay=date(2026, 3, 2), dia_danh="Cần Thơ",
                noi_dung="Nội dung.", noi_nhan=["Như trên", "Lưu: VT"], nguoi_ky="Nguyễn Văn A", chuc_vu_ky="GIÁM ĐỐC")
    base.update(kw)
    return CongVan(**base)


def _txt(data: bytes) -> str:
    d = Document(io.BytesIO(data))
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        out += [c.text for r in t.rows for c in r.cells]
    return "\n".join(out)


def _png_b64(nen="white") -> str:
    im = Image.new("RGB", (200, 100), nen)
    ImageDraw.Draw(im).ellipse((20, 20, 180, 80), outline="blue", width=6)
    b = io.BytesIO()
    im.save(b, "PNG")
    return base64.b64encode(b.getvalue()).decode()


# ---------------------------------------------------------------- thành phần thể thức
def test_bo_quoc_hieu_tieu_ngu():
    out, _ = cong_van_to_bytes_ex(_cv(thanh_phan={"quoc_hieu": False, "tieu_ngu": False}))
    t = _txt(out)
    assert "CỘNG HÒA" not in t and "Độc lập" not in t
    assert "SỞ A" in t and "GIÁM ĐỐC" in t
    out2, _ = cong_van_to_bytes_ex(_cv())
    assert "CỘNG HÒA" in _txt(out2) and "Độc lập" in _txt(out2)


def test_profile_dang_va_chuong():
    out, _ = cong_van_to_bytes_ex(_cv(co_quan_ban_hanh="ĐẢNG ỦY X"), "hd05_dang")
    assert "ĐẢNG CỘNG SẢN VIỆT NAM" in _txt(out)


def test_tuy_chinh_hop_le_va_khong_hop_le():
    p = ap_dung_tuy_chinh(load_profile("nd30"), {"body": {"size_pt": 13}})
    assert p["body"]["size_pt"] == 13
    with pytest.raises(ValueError):
        ap_dung_tuy_chinh(load_profile("nd30"), {"khong_co": 1})
    with pytest.raises(ValueError):
        ap_dung_tuy_chinh(load_profile("nd30"), {"body": {"size_pt": 9999}})


def test_chu_ky_va_con_dau_dung_docx():
    out, _ = cong_van_to_bytes_ex(_cv(chu_ky_anh={"b64": _png_b64(), "rong_mm": 38},
                                      con_dau_anh={"b64": _png_b64(), "duong_kinh_mm": 40}))
    d = Document(io.BytesIO(out))
    assert len(d.inline_shapes) >= 1                          # chữ ký nằm cùng dòng
    assert "wp:anchor" in d.element.xml                        # con dấu nổi


# ---------------------------------------------------------------- ký tự lạ
def test_quet_ky_tu_la():
    s = "Xin chào​ bạn **đậm** [oaicite:1]{index=1} 😀 tôi"
    ma = {p.ma for p in ky_tu_la.quet(s)}
    assert len(ma) >= 3
    sach, tk = ky_tu_la.lam_sach(s)
    assert "​" not in sach and "oaicite" not in sach and "**" not in sach
    assert tk


def test_quet_van_ban_sach_khong_bao_gia():
    assert ky_tu_la.quet("Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015.") == []


# ---------------------------------------------------------------- khung VBN1
def test_khung_khu_hoi():
    g = khung.dong_goi({"a": 1}, [("x", b"\x00\x01"), ("y", b"abc")])
    assert khung.la_khung(g)
    meta, tep = khung.tach(g)
    assert meta["a"] == 1 and tep["x"] == b"\x00\x01" and tep["y"] == b"abc"


# ---------------------------------------------------------------- slide
def _slides():
    return [SlideIn(title="Doanh thu quý 3 tăng 24% nhờ kênh bán trực tuyến",
                    items=[(0, "Doanh thu đạt 120 tỷ đồng", True), (0, "Chi phí giảm 8%", True),
                           (0, "Khách mới tăng 3 lần", True)]),
            SlideIn(title="Tổng quan", items=[(0, f"Bước {i} cần làm", True) for i in range(1, 8)])]


@pytest.mark.parametrize("pf", ["slide_action", "slide_bao_cao", "slide_thuyet_trinh", "slide_ban_hang",
                                "slide_chuan", "slide_hoc_thuat"])
def test_dung_slide_moi_muc_dich(pf):
    out, _, n = dung_slide(_slides(), ThongTin(giang_vien="A", don_vi="B"), profile=pf)
    assert n >= 3 and out[:2] == b"PK"


def test_tieu_de_moi_va_tuy_chinh():
    from pptx import Presentation
    out, _, _ = dung_slide(_slides(), ThongTin(giang_vien="A", don_vi="B"), "slide_action",
                           tuy_chinh={"body": {"size_pt": 20}}, tieu_de_moi={"2": "Bảy bước giúp ta hoàn thành kế hoạch"})
    ts = [s.shapes.title.text_frame.text for s in Presentation(io.BytesIO(out)).slides if s.shapes.title]
    assert any("Bảy bước giúp ta" in t for t in ts)
    with pytest.raises(ValueError):
        dung_slide(_slides(), ThongTin(), "slide_action", tuy_chinh={"khong_co": 1})


def test_lint_mau_va_slide():
    p = load_profile("slide_action")
    kq = slide_lint.phan_tich(_slides(), p)
    assert any(d["ma"] == "action_title" and d["slide"] == 2 for d in kq["phat_hien"])
    assert any(d["ma"] == "chunking" for d in kq["phat_hien"])
    p["colors"]["text"] = "CCCCCC"
    assert any(d["ma"] == "tuong_phan" and d["muc"] == "canh_bao" for d in slide_lint.kiem_tra_mau(p))
    assert slide_lint.ti_le_tuong_phan("000000", "FFFFFF") == pytest.approx(21, abs=0.01)


# ---------------------------------------------------------------- PDF
def _pdf_mau(n=3) -> bytes:
    from pypdf import PdfWriter
    doc = Document()
    doc.add_paragraph("CÔNG VĂN").runs[0].bold = True
    for i in range(n):
        doc.add_paragraph(f"Nội dung trang {i + 1}")
        if i < n - 1:
            doc.add_page_break()
    b = io.BytesIO()
    doc.save(b)
    return b.getvalue()


def _pdf_that(n=3) -> bytes:
    from pypdf import PdfWriter
    w = PdfWriter()
    for _ in range(n):
        w.add_blank_page(595, 842)
    b = io.BytesIO()
    w.write(b)
    return b.getvalue()


def test_pdf_watermark_bookmark_dau():
    from pypdf import PdfReader
    from app import pdf_tools
    out, canh, info = pdf_tools.xu_ly_pdf(_pdf_that(), {
        "bookmark": {"giu_goc": False, "danh_sach": [{"tieu_de": "Phần 1", "trang": 1, "cap": 0},
                                                     {"tieu_de": "Phần 2", "trang": 3, "cap": 0}]},
        "watermark_chu": {"bat": True, "chu": "BẢN NHÁP", "do_dam": 30, "goc": 45},
        "chu_ky": {"bat": True, "trang": "cuoi", "x_pct": 70, "y_pct": 80, "rong_mm": 40},
        "thong_tin": {"tieu_de": "Thử"},
    }, {"ky_anh": base64.b64decode(_png_b64())})
    r = PdfReader(io.BytesIO(out))
    assert len(r.pages) == 3
    assert [o.title for o in r.outline] == ["Phần 1", "Phần 2"]
    assert len(out) > len(_pdf_that())


# ---------------------------------------------------------------- lớp API
def test_api_thong_tin_va_profile():
    r = api_core.xu_ly("GET", "/api/thong-tin", {})
    d = json.loads(r.body)
    assert {x["id"] for x in d["slide"]} >= {"slide_action", "slide_ban_hang"}
    assert {x["id"] for x in d["the_thuc"]} >= {"nd30", "hd05_dang"}
    assert api_core.xu_ly("GET", "/api/profile/khong_co", {}).status == 404
    assert api_core.xu_ly("GET", "/api/khong-co", {}).status == 404


def test_api_cong_van_va_canh_bao_header():
    body = _cv(thanh_phan={"quoc_hieu": False}).model_dump_json().encode()
    r = api_core.xu_ly("POST", "/api/cong-van/docx", {}, body)
    assert r.status == 200 and "CỘNG HÒA" not in _txt(r.body)
    assert api_core.xu_ly("POST", "/api/cong-van/docx", {}, b"{}").status == 422


def test_api_ky_tu_la():
    r = api_core.xu_ly("POST", "/api/ky-tu-la", {}, json.dumps({"text": "a​b **x**", "lam_sach": True}).encode())
    d = json.loads(r.body)
    assert d["phat_hien"] and "​" not in d["van_ban_sach"]


def test_api_import_docx_co_ky_tu_la():
    doc = Document()
    doc.add_paragraph("SỞ A")
    doc.add_paragraph("Nội dung​ có ký tự lạ [oaicite:1]{index=1}")
    b = io.BytesIO()
    doc.save(b)
    r = api_core.xu_ly("POST", "/api/import", {"ten": "a.docx"}, b.getvalue())
    d = json.loads(r.body)
    assert r.status == 200 and d["ky_tu_la"]


def test_api_pdf_phan_tich_va_xu_ly():
    r = api_core.xu_ly("POST", "/api/pdf/phan-tich", {}, _pdf_that())
    assert json.loads(r.body)["so_trang"] == 3
    g = khung.dong_goi({"tuy_chon": {"watermark_chu": {"bat": True, "chu": "MẬT", "do_dam": 20}}}, [("pdf", _pdf_that())])
    r = api_core.xu_ly("POST", "/api/pdf/xu-ly", {}, g)
    assert r.status == 200 and r.body[:4] == b"%PDF"
    assert api_core.xu_ly("POST", "/api/pdf/xu-ly", {}, b"x").status == 400


def test_api_slide_phan_tich_va_chuan_hoa():
    out, _, _ = dung_slide(_slides(), ThongTin(giang_vien="A", don_vi="B"), "slide_action")
    r = api_core.xu_ly("POST", "/api/slide/phan-tich", {"ten": "a.pptx", "kieu": "slide_ban_hang"}, out)
    assert r.status == 200 and "phat_hien" in json.loads(r.body)
    g = khung.dong_goi({"ten": "a.pptx", "kieu": "slide_thuyet_trinh", "giang_vien": "A", "don_vi": "B",
                        "tuy_chinh": {"body": {"size_pt": 24}}, "tieu_de_moi": {"1": "Tiêu đề mới giúp rõ ý"}},
                       [("pptx", out)])
    r = api_core.xu_ly("POST", "/api/slide/chuan-hoa", {}, g)
    assert r.status == 200 and r.body[:2] == b"PK"
    bad = api_core.xu_ly("POST", "/api/slide/chuan-hoa", {"ten": "a.pptx", "tuy_chinh": "{x"}, out)
    assert bad.status == 400


def test_fastapi_catch_all():
    from fastapi.testclient import TestClient
    from app.main import app
    c = TestClient(app)
    assert c.get("/api/tac-gia").status_code == 200
    assert "Thevancorp-1801808400" in c.get("/api/tac-gia").headers["x-tac-gia"]
    r = c.post("/api/cong-van/docx", content=_cv().model_dump_json())
    assert r.status_code == 200 and r.content[:2] == b"PK"
    assert c.get("/").status_code == 200


# ---------------------------------------------------------------- thể thức mới
@pytest.mark.parametrize("pf", ["tt92_bqp", "qd933_cong_doan", "hd29_doan"])
def test_the_thuc_moi_dung_duoc(pf):
    out, _ = cong_van_to_bytes_ex(_cv(), pf)
    assert Document(io.BytesIO(out)).paragraphs is not None
    p = load_profile(pf)
    assert p["loai"] == "van_ban" and p["thanh_phan"]["quoc_hieu"] is True


def test_hd29_doan_tieu_de_va_dau_sao():
    out, _ = cong_van_to_bytes_ex(_cv(co_quan_chu_quan="BAN CHẤP HÀNH TRUNG ƯƠNG"), "hd29_doan")
    t = _txt(out)
    assert "ĐOÀN TNCS HỒ CHÍ MINH" in t and "***" in t and "CỘNG HÒA" not in t and "Độc lập" not in t


def test_tt92_le_quan_doi():
    assert load_profile("tt92_bqp")["margins_mm"] == {"top": 25, "bottom": 20, "left": 35, "right": 15}


# ---------------------------------------------------------------- hàng loạt theo chế độ
def _lo():
    return json.loads(api_core.xu_ly("POST", "/api/lo/tao", {}).body)["lo"]


def _docx_vb() -> bytes:
    d = Document()
    d.add_paragraph("SỞ A")
    d.add_paragraph("Nội dung​ thử [oaicite:1]{index=1}")
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


def test_lo_che_do_word_bo_qua_file_khac_loai():
    lo = _lo()
    ok = json.loads(api_core.xu_ly("POST", f"/api/lo/{lo}/them", {"ten": "a.docx", "che_do": "van_ban",
                    "profile_van_ban": "tt92_bqp", "lam_sach": "1",
                    "thanh_phan": json.dumps({"quoc_hieu": False, "tieu_ngu": False})}, _docx_vb()).body)
    assert ok["loai"] == "Word" and ok["trang_thai"] in ("ok", "can_ra_soat")
    bo = json.loads(api_core.xu_ly("POST", f"/api/lo/{lo}/them", {"ten": "b.pdf", "che_do": "van_ban"}, _pdf_that()).body)
    assert bo["trang_thai"] == "bo_qua"
    from app import lo as lo_mod
    import zipfile
    z = zipfile.ZipFile(lo_mod.tao_zip(lo))
    ten = [n for n in z.namelist() if n.endswith(".docx")]
    assert len(ten) == 1
    assert "CỘNG HÒA" not in _txt(z.read(ten[0]))
    assert "Bỏ qua" in z.read("BAO-CAO.csv").decode("utf-8-sig")


def test_lo_che_do_pdf_khung():
    lo = _lo()
    o = {"bookmark": {"giu_goc": True, "tu_dong": True}, "watermark_chu": {"bat": True, "chu": "NHÁP", "do_dam": 20}}
    g = khung.dong_goi({"ten": "x/a.pdf", "che_do": "pdf", "tuy_chon_pdf": o}, [("tep", _pdf_that())])
    r = json.loads(api_core.xu_ly("POST", f"/api/lo/{lo}/them", {}, g).body)
    assert r["loai"] == "PDF" and r["ra"].endswith(".pdf") and r["thong_ke"] == "3 trang", r


def test_lo_che_do_slide():
    out, _, _ = dung_slide(_slides(), ThongTin(giang_vien="A"), "slide_action")
    lo = _lo()
    r = json.loads(api_core.xu_ly("POST", f"/api/lo/{lo}/them", {"ten": "s.pptx", "che_do": "slide", "kieu": "slide_ban_hang"}, out).body)
    assert r["loai"] == "Slide" and r["ra"].endswith(".pptx")
    bo = json.loads(api_core.xu_ly("POST", f"/api/lo/{lo}/them", {"ten": "s.docx", "che_do": "slide"}, _docx_vb()).body)
    assert bo["trang_thai"] == "bo_qua"
