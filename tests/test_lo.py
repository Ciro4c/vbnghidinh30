# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra xử lý hàng loạt. Chạy:  python -m pytest -q"""
import csv
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import _mau  # noqa: E402
from app import lo  # noqa: E402


def test_sach_duong_dan():
    assert lo.sach_duong_dan("../../etc/passwd") == "etc/passwd"
    assert lo.sach_duong_dan("C:\\a\\b:c.docx") == "C_/a/b_c.docx"
    assert lo.sach_duong_dan("") == "tep"


def test_lo_nhieu_file_va_zip():
    ma = lo.tao_lo()
    docx, pptx = _mau.docx_co_bang_hinh(), _mau.pptx_co_bang_hinh()
    a = lo.xu_ly_tep(ma, "Thu muc/Van ban 1.docx", docx)
    b = lo.xu_ly_tep(ma, "Thu muc/Van ban 1.docx", docx)           # trùng tên: không ghi đè
    c = lo.xu_ly_tep(ma, "bai giang.pptx", pptx, {"giang_vien": "A", "don_vi": "B"})
    d = lo.xu_ly_tep(ma, "cu.doc", b"x")
    e = lo.xu_ly_tep(ma, "hong.docx", b"khong phai docx")
    f = lo.xu_ly_tep(ma, "ghi-chu.pdf", b"x")
    assert a["trang_thai"] in ("ok", "can_ra_soat") and a["thong_ke"] == "1 bảng, 1 hình"
    assert a["ra"] != b["ra"]
    assert c["trang_thai"] in ("ok", "can_ra_soat") and c["loai"] == "Slide"
    assert d["trang_thai"] == "loi" and "Lưu thành" in d["canh_bao"][0]
    assert e["trang_thai"] == "loi" and f["trang_thai"] == "loi"
    duong = lo.tao_zip(ma)
    with zipfile.ZipFile(duong) as z:
        ten = z.namelist()
        assert "BAO-CAO.csv" in ten
        assert sum(n.startswith("ket-qua/") for n in ten) == 3
        rows = list(csv.reader(io.StringIO(z.read("BAO-CAO.csv").decode("utf-8-sig"))))
    assert len(rows) == 7 and rows[0][0] == "STT"


def test_ma_lo_khong_hop_le():
    try:
        lo.xu_ly_tep("../x", "a.docx", b"")
        raise AssertionError("phải báo lỗi")
    except ValueError:
        pass
