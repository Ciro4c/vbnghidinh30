# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Schema dữ liệu đầu vào cho từng loại văn bản."""
from __future__ import annotations

from datetime import date

from pydantic import BaseModel, Field, field_validator


# khóa -> tên loại in trên văn bản (None = công văn, không có tên loại)
TEN_LOAI = {
    "cong_van": None,
    "quyet_dinh": "QUYẾT ĐỊNH",
    "to_trinh": "TỜ TRÌNH",
    "thong_bao": "THÔNG BÁO",
    "bao_cao": "BÁO CÁO",
    "ke_hoach": "KẾ HOẠCH",
    "bien_ban": "BIÊN BẢN",
    "chi_thi": "CHỈ THỊ",
    "huong_dan": "HƯỚNG DẪN",
    "chuong_trinh": "CHƯƠNG TRÌNH",
    "quy_che": "QUY CHẾ",
    "quy_dinh": "QUY ĐỊNH",
    "nghi_quyet": "NGHỊ QUYẾT",
    "giay_moi": "GIẤY MỜI",
    "giay_gioi_thieu": "GIẤY GIỚI THIỆU",
    "de_an": "ĐỀ ÁN",
    "phuong_an": "PHƯƠNG ÁN",
}


def _kiem_tra_loai(v):
    if v not in TEN_LOAI:
        raise ValueError(f"Loại văn bản không hợp lệ: {v!r}")
    return v


def _tach_dong(v):
    if isinstance(v, str):
        return [s for s in (x.strip() for x in v.splitlines()) if s]
    return v


class CongVan(BaseModel):
    loai_van_ban: str = "cong_van"                   # khóa trong TEN_LOAI
    co_quan_chu_quan: str = ""                       # vd: UBND TỈNH ABC (có thể để trống)
    co_quan_ban_hanh: str = Field(min_length=1)      # vd: SỞ GIÁO DỤC VÀ ĐÀO TẠO
    so: str = ""                                     # vd: 123
    ky_hieu: str = ""                                # vd: SGDĐT-VP
    dia_danh: str = ""                               # vd: Cần Thơ
    ngay: date
    trich_yeu: str = ""                              # nội dung sau "V/v"
    kinh_gui: list[str] = Field(default_factory=list)
    noi_dung: str = Field(min_length=1)              # dùng ký hiệu: "# ", "## ", "- "
    quyen_han_ky: str = ""                           # vd: KT. GIÁM ĐỐC
    chuc_vu_ky: str = ""                             # vd: PHÓ GIÁM ĐỐC
    nguoi_ky: str = ""
    noi_nhan: list[str] = Field(default_factory=list)
    phu_luc: str = ""                                # nội dung phụ lục sau chữ ký (cùng ký hiệu như noi_dung)
    phu_luc_ngang: bool = False                      # phụ lục in trên giấy A4 nằm ngang (bảng rộng)
    hien_so_trang_dau: bool = False                  # NĐ 30: không hiện số ở trang 1
    # Hình trong nội dung: {"1": {"mime","b64","w_mm","w_px","h_px"}}, nội dung tham chiếu bằng dòng [[hinh:1]]
    hinh: dict[str, dict] = Field(default_factory=dict)

    _loai = field_validator("loai_van_ban")(_kiem_tra_loai)
    _tach = field_validator("kinh_gui", "noi_nhan", mode="before")(_tach_dong)
