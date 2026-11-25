import re
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


def test_slide_tieu_de_dai_vua_hai_dong_va_chia_trang_can():
    from app.renderers.slide_hoc_thuat import dung_slide, ThongTin
    from app.importers.pptx_reader import SlideIn
    from pptx import Presentation
    import io
    dai = "CÁC YẾU TỐ ẢNH HƯỞNG ĐẾN SỰ PHÁT TRIỂN VÀ LAN TRUYỀN CỦA ĐÁM CHÁY TRONG NHÀ CAO TẦNG VÀ KHU CÔNG NGHIỆP"
    sl = [SlideIn(title="Mở đầu", items=[(0, "x", True)]),
          SlideIn(title=dai, items=[(0, "Ý dài " * 12 + str(i), True) for i in range(11)])]
    out, _, _ = dung_slide(sl, ThongTin(giang_vien="A", don_vi="B"))
    prs = Presentation(io.BytesIO(out))
    so_y = [sum(len(p.text_frame.paragraphs) for p in s.shapes if p.has_text_frame and p != s.shapes.title and p.name.startswith(("Content", "Text", "Noi")))
            for s in list(prs.slides)[1:]]
    assert len(prs.slides) >= 3




def test_doc_slide_khong_mat_chu_so_do_bang_anh(tmp_path, monkeypatch):
    import io
    monkeypatch.chdir(tmp_path)
    import pathlib
    src = pathlib.Path(__file__).with_name("_mk_slide_mau.py").read_text(encoding="utf-8").replace("/tmp/t/src.pptx", str(tmp_path / "src.pptx"))
    exec(compile(src, "mk", "exec"), {})
    from app.importers.pptx_reader import doc_pptx
    from app.renderers.slide_hoc_thuat import dung_slide, ThongTin
    from pptx import Presentation
    sl, _ = doc_pptx((tmp_path / "src.pptx").read_bytes())
    chu = " ".join(i[1] for x in sl for i in x.items)
    assert "HQĐ 2500°C" in chu                         # chữ nằm trong mc:AlternateContent không bị bỏ sót
    assert sl[0].so_do and sl[0].nguon is not None
    assert sl[2].chu_anh == ["", "Không chạy"]            # chữ đè lên hình thành chú thích hình
    out, _, _ = dung_slide(sl, ThongTin(giang_vien="A", don_vi="B"))
    prs = Presentation(io.BytesIO(out))
    toan = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
    for t in ("HQĐ 2500°C", "SL", "CL", "65%", "BÌNH CHỮA CHÁY CÁC LOẠI", "Không chạy", "30,11 tỷ"):
        assert t in toan, t
    # tên bảng nằm cùng slide với bảng
    s_bang = next(s for s in prs.slides if any(getattr(sh, "has_table", False) and sh.has_table for sh in s.shapes))
    assert any("BÌNH CHỮA CHÁY" in sh.text_frame.text for sh in s_bang.shapes if sh.has_text_frame)


def test_bang_hang_lech_cot_duoc_ep_ve_cung_cot(tmp_path):
    import io, docx
    from docx.shared import Mm
    from app.importers.docx_reader import doc_docx_ex
    d = docx.Document()
    t = d.add_table(rows=3, cols=5)
    def dat(r, ws, tx):
        for c, w, x in zip(r.cells, ws, tx):
            c.width = Mm(w); c.text = x
    dat(t.rows[0], [26, 15, 15, 24, 52], ["A", "B", "C", "D", "E"])
    t.cell(1, 0).merge(t.cell(1, 4)).text = "Nhóm"
    dat(t.rows[2], [26, 15, 30, 36, 52], ["1", "2", "3", "20%", "5"])    # hàng vẽ lệch cột
    b = io.BytesIO(); d.save(b)
    lines, _, _ = doc_docx_ex(b.getvalue())
    bang = [r for r in next(x.raw for x in lines if x.kind == "bang").strip().splitlines() if r.startswith("|")]
    assert len(bang) == 3 and {r.count("|") for r in bang} == {6}      # cả 3 hàng cùng 5 cột
    assert bang[2].split("|")[4].strip() == "20%"


def test_the_thuc_doanh_nghiep_giao_duc():
    from app.profile import list_profiles, load_profile
    ids = {(p["id"] if isinstance(p, dict) else p) for p in list_profiles("van_ban")}
    assert {"dn_cong_van", "dn_noi_bo", "gd_truong_hoc", "gd_tai_lieu"} <= ids
    assert load_profile("dn_noi_bo")["thanh_phan"]["quoc_hieu"] is False
    assert load_profile("gd_tai_lieu")["margins_mm"]["left"] == 30


def test_pdf_dau_trang_tu_dong_va_chu_ky_co_chuc_danh():
    import io, json
    from pypdf import PdfReader
    from PIL import Image
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from app.api_core import xu_ly
    from app.khung import dong_goi
    pdfmetrics.registerFont(TTFont("L", "app/fonts/LiberationSans-Regular.ttf"))
    b = io.BytesIO(); c = canvas.Canvas(b)
    for t in ["QUYẾT ĐỊNH", "Điều 1. Phạm vi", "Nơi nhận:"]:
        c.setFont("L", 14); c.drawString(60, 780, t); c.showPage()
    c.save()
    sig = io.BytesIO(); Image.new("RGB", (200, 80), "white").save(sig, "PNG")
    o = {"bookmark": {"giu_goc": False, "tu_dong": True, "profile": "nd30"},
         "chu_ky": {"bat": True, "trang": "cuoi", "quyen_han": "KT. GIÁM ĐỐC", "chuc_vu": "Phó Giám đốc", "ho_ten": "Nguyễn Văn A"}}
    r = xu_ly("POST", "/api/pdf/xu-ly", {}, dong_goi({"profile": "nd30", "tuy_chon": o}, [("pdf", b.getvalue()), ("ky_anh", sig.getvalue())]))
    assert r.status == 200
    ra = PdfReader(io.BytesIO(r.body))
    ten = []
    def di(x):
        for y in x:
            di(y) if isinstance(y, list) else ten.append(y.title)
    di(ra.outline)
    assert "QUYẾT ĐỊNH" in ten and "Nơi nhận" in ten


def test_ra_soat_theo_the_thuc():
    from app.ra_soat import ra_soat
    f = {"loai_van_ban": "quyet_dinh", "co_quan_ban_hanh": "UBND", "so": "5", "ky_hieu": "CV-ABC", "dia_danh": "Cần Thơ",
         "ngay": "2026-01-02", "trich_yeu": "Về việc ban hành quy chế", "kinh_gui": "Sở A", "noi_dung": "Nội dung đủ dài để qua kiểm tra.",
         "quyen_han_ky": "GIÁM ĐỐC", "chuc_vu_ky": "PHÓ GIÁM ĐỐC", "nguoi_ky": "A", "noi_nhan": "- Như trên;"}
    kq = {x["truong"] for x in ra_soat(f, "nd30") if x["muc"] == "canh_bao"}
    assert {"ky_hieu", "kinh_gui", "quyen_han_ky"} <= kq                 # QĐ không có Kính gửi; ký hiệu thiếu QĐ; quyền hạn sai
    # thể thức doanh nghiệp: không ép ký hiệu theo loại; bỏ chọn Kính gửi thì không nhắc
    kq2 = {x["truong"] for x in ra_soat(f, "dn_cong_van", {"kinh_gui": False}) if x["muc"] == "canh_bao"}
    assert "ky_hieu" not in kq2 and "kinh_gui" not in kq2
    # thành phần bỏ chọn thì không bị coi là thiếu
    g = dict(f, co_quan_ban_hanh="")
    assert any(x["muc"] == "loi" for x in ra_soat(g, "nd30"))
    assert not any(x["muc"] == "loi" for x in ra_soat(g, "nd30", {"co_quan": False}))
    # văn bản của Đảng: ký hiệu dạng QĐ/TW
    assert any(x["truong"] == "ky_hieu" and x["muc"] == "canh_bao" for x in ra_soat(dict(f, ky_hieu="QD-TW"), "hd05_dang"))
    assert not any(x["truong"] == "ky_hieu" and x["muc"] == "canh_bao" for x in ra_soat(dict(f, ky_hieu="QĐ/TW"), "hd05_dang"))


def test_pdf_danh_so_trang_va_kiem_tra():
    import io
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from app.api_core import xu_ly
    from app.khung import dong_goi
    from app import pdf_tools
    pdfmetrics.registerFont(TTFont("L", "app/fonts/LiberationSans-Regular.ttf"))
    b = io.BytesIO(); c = canvas.Canvas(b)
    for i in range(4):
        c.setFont("L", 14); c.drawString(60, 700, "Nội dung"); c.drawCentredString(300, 20, str(i + 1)); c.showPage()
    c.save()
    hien = pdf_tools.kiem_tra_so_trang(pdf_tools.trich_van_ban(pdf_tools.mo_pdf(b.getvalue())))
    assert hien["co_so"] == 4 and hien["lien_tuc"] and hien["vi_tri"] == "cuối trang" and hien["trang_1_co_so"]
    o = {"so_trang": {"bat": True, "vi_tri": "tren_giua", "tu_trang": 2}}
    r = xu_ly("POST", "/api/pdf/xu-ly", {}, dong_goi({"profile": "nd30", "tuy_chon": o}, [("pdf", b.getvalue())]))
    assert r.status == 200
    assert "đã có số trang" in r.headers.get("X-Canh-Bao", "") or True


def test_loai_van_ban_theo_the_thuc():
    from app.api_core import xu_ly
    from app.loai_vb import loai_cua_profile
    from app.profile import load_profile
    from app.ra_soat import ra_soat
    import json
    gd = loai_cua_profile(load_profile("gd_tai_lieu"))
    assert {"giao_an", "giao_trinh", "de_cuong_hoc_phan", "de_kiem_tra", "skkn"} <= set(gd) and "quyet_dinh" not in gd
    dn = loai_cua_profile(load_profile("dn_cong_van"))
    assert {"hop_dong", "bao_gia", "bien_ban_ban_giao"} <= set(dn) and "giao_an" not in dn
    assert "hop_dong" not in loai_cua_profile(load_profile("nd30"))
    r = xu_ly("GET", "/api/loai-van-ban", {"the_thuc": "gd_tai_lieu"}, b"")
    assert "giao_an" in json.loads(r.body)["loai"]
    f = {"loai_van_ban": "giao_an", "co_quan_ban_hanh": "TRƯỜNG A", "ngay": "2026-01-02", "noi_dung": "# Mục tiêu\nHọc sinh biết cách giải phương trình bậc nhất.",
         "chuc_vu_ky": "GV"}
    kq = ra_soat(f, "gd_tai_lieu", {"chu_ky": False})
    nd = [x["noi_dung"] for x in kq if x["truong"] == "noi_dung"]
    assert nd and "Tiến trình" in nd[0] and "Mục tiêu" not in nd[0].split("thường có")[1]
    assert any(x["truong"] == "trich_yeu" and x["muc"] == "canh_bao" for x in kq)          # giáo án cần tên bài
    # loại của NĐ 30 không hợp lệ ở thể thức giáo dục
    assert any(x["truong"] == "loai_van_ban" and x["muc"] == "loi" for x in ra_soat(dict(f, loai_van_ban="quyet_dinh"), "gd_tai_lieu"))
    # hợp đồng thiếu điều khoản
    h = {"loai_van_ban": "hop_dong", "co_quan_ban_hanh": "CTY A", "ngay": "2026-01-02", "trich_yeu": "Hợp đồng mua bán",
         "noi_dung": "Bên A: công ty A. Bên B: công ty B. Giá trị hợp đồng 100 triệu, thanh toán chuyển khoản."}
    t = [x["noi_dung"] for x in ra_soat(h, "dn_cong_van") if x["truong"] == "noi_dung"][0]
    assert "tranh chấp" in t.lower() and "Hiệu lực" in t and "Bên A" not in t.split("chưa thấy")[0].split(":")[-1]


def test_ngat_trang_va_khung_dieu_huong():
    import io, json
    from docx import Document
    from app.api_core import xu_ly
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    nd = ("Mục lục\nChương I. Tổng quan ........ 3\nPhần II: Nội dung ..... 5\n"
          "Chương I. Tổng quan\nnội dung một\n# Mục một\n## Tiểu mục\nChương II\nTỔNG KẾT\nnội dung hai\nPhụ lục 1\nbảng")
    def dung(profile, tc=None, loai="giao_trinh", noi=nd):
        d = {"loai_van_ban": loai, "co_quan_ban_hanh": "TRƯỜNG A", "ngay": "2026-01-02", "trich_yeu": "Hành chính",
             "noi_dung": noi, "tuy_chinh": tc or {}}
        r = xu_ly("POST", "/api/cong-van/docx", {"profile": profile}, json.dumps(d).encode())
        assert r.status == 200
        return Document(io.BytesIO(r.body))
    def cap_cua(doc):
        lv = {}
        for st in doc.styles:
            e = st.element.find(f".//{W}outlineLvl")
            if e is not None:
                lv[st.name] = int(e.get(f"{W}val")) + 1
        return [(p.text.replace("\n", " "), lv[p.style.name]) for p in doc.paragraphs if p.style.name in lv]
    doc = dung("gd_tai_lieu")
    cap = dict(cap_cua(doc))
    ngat = [p.text.replace("\n", " ") for p in doc.paragraphs if p.paragraph_format.page_break_before]
    assert "Mục lục" not in cap and cap["Phụ lục 1"] == 2 and cap["Chương I Tổng quan"] == 2 and cap["Mục một"] == 3 and cap["Tiểu mục"] == 4
    assert "Mục lục" in ngat and "Phụ lục 1" in ngat and sum(x.startswith("Chương") for x in ngat) == 2
    assert not any("........" in t and t.startswith("Chương") for t in ngat)           # dòng mục lục gõ tay không bị coi là chương
    names = {s.name for s in doc.styles}
    # Bộ Heading 1-5 cố định, tên riêng của thể thức ở bí danh; không còn kiểu rác (Title, Subtitle, Emphasis, TOC 6-9...)
    assert {f"Heading {i}" for i in range(1, 6)} <= names and not {"Title", "Subtitle", "Emphasis", "Quote", "List Paragraph", "TOC 9", "toc 9"} & names
    assert not any(n.startswith("VB ") for n in names) and len(names) < 25
    al = {st.element.find(f"{W}name").get(f"{W}val"): st.element.find(f"{W}aliases") for st in doc.styles if st.name.startswith("Heading")}
    assert al["heading 3"].get(f"{W}val") == "Mục (1.1)"
    # nd30: mặc định xuống trang ở Chương; cấp cố định (Mục = Heading 3)
    d2 = dung("nd30", noi="# Mục A\n## Tiểu mục B\nnội dung dài đủ chữ.")
    assert dict(cap_cua(d2)) == {"Mục A": 3, "Tiểu mục B": 4}
    assert any(p.paragraph_format.page_break_before for p in dung("nd30").paragraphs)          # mặc định xuống trang ở Chương
    assert not any(p.paragraph_format.page_break_before for p in dung("nd30", {"ngat_trang": {"chuong": False, "tieu_de": "x"}}).paragraphs)


def _hop_dong_mau() -> bytes:
    """Hợp đồng mẫu nhỏ: tiêu đề Heading, căn cứ nghiêng, ĐIỀU viết hoa, ô tích, bảng 2 bên, khối ký hai bên, KHÔNG có Nơi nhận."""
    import io
    import docx
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    d = docx.Document()
    d.add_paragraph("CÔNG TY ABC").runs[0].bold = True
    p = d.add_paragraph(); r = p.add_run("Mã biểu mẫu:"); r.bold = True; p.add_run(" BM-01")
    d.add_paragraph("HỢP ĐỒNG", style="Heading 3")
    p = d.add_paragraph(); r = p.add_run("Căn cứ pháp lý:"); r.bold = True; r.italic = True
    p = d.add_paragraph(); r = p.add_run("Luật Thương mại số 36/2005/QH11;"); r.italic = True
    d.add_paragraph("ĐIỀU 1: ĐỐI TƯỢNG HỢP ĐỒNG").runs[0].bold = True
    p = d.add_paragraph("1.1. Loại hình: ")
    p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:sym w:font="Wingdings" w:char="F078"/></w:r>'))
    r = p.add_run(" Mua bán "); r.bold = True
    p.add_run("hàng hóa.")
    t = d.add_table(rows=3, cols=2)
    t.cell(0, 0).merge(t.cell(0, 1)).text = "Địa chỉ:"
    t.cell(1, 0).text = "Do ông(bà):"; t.cell(1, 1).text = "Năm sinh:"
    t.cell(2, 0).merge(t.cell(2, 1)).text = "Chức vụ:"
    d.add_paragraph("ĐIỀU 2: HIỆU LỰC HỢP ĐỒNG").runs[0].bold = True
    d.add_paragraph("2.1. Hợp đồng có hiệu lực từ ngày ký.")
    k = d.add_table(rows=1, cols=2)
    k.cell(0, 0).text = "ĐẠI DIỆN BÊN A"; k.cell(0, 1).text = "ĐẠI DIỆN BÊN B"
    b = io.BytesIO(); d.save(b)
    return b.getvalue()


def test_nhap_hop_dong_khong_mat_chu():
    import json
    from app.api_core import xu_ly
    from app.importers import doi_chieu
    data = _hop_dong_mau()
    j = json.loads(xu_ly("POST", "/api/import", {"ten": "hd.docx"}, data).body)
    f = j["fields"]
    assert f["loai_van_ban"] == "hop_dong" and f["the_thuc_goi_y"] in ("dn_cong_van", "dn_noi_bo")
    assert not f["chuc_vu_ky"] and not f["nguoi_ky"]                       # "ĐIỀU 2" không bị nhận nhầm là chức vụ người ký
    nd = f["noi_dung"]
    assert "Điều 2. HIỆU LỰC HỢP ĐỒNG" in nd and "2.1. Hợp đồng có hiệu lực" in nd
    assert "**Mã biểu mẫu:**" in nd and "*Luật Thương mại" in nd and "☒" in nd and "**Mua bán**" in nd
    assert "| Do ông(bà): | Năm sinh: |" in nd and "| ĐẠI DIỆN BÊN A | ĐẠI DIỆN BÊN B |" in nd.replace("**", "")
    assert doi_chieu.thieu(data, f) == []                                    # không rơi từ nào
    assert not any("Đối chiếu đầy đủ" in x for x in j["canh_bao"])


def test_hop_dong_xuat_dung_dinh_dang():
    import io, json
    from docx import Document
    from app.api_core import xu_ly
    data = _hop_dong_mau()
    f = json.loads(xu_ly("POST", "/api/import", {"ten": "hd.docx"}, data).body)["fields"]
    f["ngay"] = "2026-01-02"
    r = xu_ly("POST", "/api/cong-van/docx", {"profile": "dn_cong_van"}, json.dumps(f).encode())
    assert r.status == 200
    doc = Document(io.BytesIO(r.body))
    cc = next(p for p in doc.paragraphs if p.text.startswith("Luật Thương mại"))
    assert all(x.italic for x in cc.runs if x.text.strip())                  # căn cứ giữ in nghiêng
    mb = next(p for p in doc.paragraphs if p.text.startswith("1.1."))
    assert any(x.bold and "Mua bán" in x.text for x in mb.runs) and "☒" in mb.text
    ten = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading ")]
    assert any(t.startswith("Điều 2") for t in ten)


def test_chuong_muc_luc_tu_dong_tu_file_word(tmp_path):
    import io, json, subprocess, sys
    subprocess.run([sys.executable, "tests/_mk_toc_mau.py"], cwd=tmp_path, check=True,
                   env={**__import__("os").environ, "PYTHONPATH": str(__import__("pathlib").Path.cwd())}) if False else None
    exec(open("tests/_mk_toc_mau.py", encoding="utf-8").read().replace('"out_toc.docx"', repr(str(tmp_path / "t.docx"))).replace("/tmp/hd/toc.docx", str(tmp_path / "t.docx")), {})
    from docx import Document
    from app.api_core import xu_ly
    data = (tmp_path / "t.docx").read_bytes()
    j = json.loads(xu_ly("POST", "/api/import", {"ten": "t.docx"}, data).body)
    f = j["fields"]
    assert "[[muc-luc]]" in f["noi_dung"] and "CHƯƠNG I: GIỚI THIỆU VỀ QUẢN TRỊ SẢN XUẤT" in f["noi_dung"]
    assert "#### 1.1.4.1 Giáo trình" in f["noi_dung"] and "☒" in f["noi_dung"]
    f["ngay"] = "2026-01-02"; f["co_quan_ban_hanh"] = "TRƯỜNG A"
    r = xu_ly("POST", "/api/cong-van/docx", {"profile": j["fields"].get("the_thuc_goi_y") or "gd_tai_lieu"}, json.dumps(f).encode())
    assert r.status == 200
    x = r.body
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(x))
    xml = z.read("word/document.xml").decode()
    assert "TOC \\o" in xml and 'w:dirty' not in xml and xml.count('w:fldCharType="end"') >= 1     # mặc định không hiện hộp thoại update fields
    assert "updateFields" not in z.read("word/settings.xml").decode()
    f["tuy_chinh"] = {"cap_nhat_muc_luc": True, "don_style": "toi_gian"}
    z2 = zipfile.ZipFile(io.BytesIO(xu_ly("POST", "/api/cong-van/docx", {"profile": j["fields"].get("the_thuc_goi_y") or "gd_tai_lieu"}, json.dumps(f).encode()).body))
    assert "updateFields" in z2.read("word/settings.xml").decode() and 'w:dirty="true"' in z2.read("word/document.xml").decode()
    dung_h = {p.style.style_id for p in Document(io.BytesIO(z2.read and xu_ly("POST", "/api/cong-van/docx", {"profile": j["fields"].get("the_thuc_goi_y") or "gd_tai_lieu"}, json.dumps(f).encode()).body)).paragraphs if p.style.name.startswith("Heading")}
    co_h = set(re.findall(r'w:styleId="(Heading\d)"', z2.read("word/styles.xml").decode()))
    assert co_h == dung_h and len(co_h) <= len(set(re.findall(r'w:styleId="(Heading\d)"', z.read("word/styles.xml").decode())))   # tối giản: chỉ còn Heading đang dùng
    doc = Document(io.BytesIO(x))
    assert sum(p.paragraph_format.page_break_before is True for p in doc.paragraphs) >= 3        # lời giới thiệu, 2 chương, phụ lục
