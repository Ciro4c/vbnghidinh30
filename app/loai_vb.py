# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Loại văn bản/tài liệu theo từng thể thức (profile).

Mỗi profile có thể khai báo mục `loai_van_ban:` (khóa -> {ten, nhan, kinh_gui, viet_tat, trich_yeu, muc, goi_y});
không khai báo thì dùng bộ loại văn bản hành chính của NĐ 30 (schemas.TEN_LOAI). Giá trị rỗng ({}) = lấy mặc định."""
from __future__ import annotations

from functools import lru_cache

from .profile import list_profiles, load_profile
from .schemas import TEN_LOAI

KHONG_KINH_GUI = {"quyet_dinh", "nghi_quyet", "nghi_dinh", "thong_tu", "thong_tu_lien_tich", "quy_che", "quy_dinh",
                  "chi_thi", "thong_cao", "ket_luan", "bien_ban"}
CO_KINH_GUI = {"cong_van", "to_trinh", "bao_cao", "de_an", "phuong_an", "giay_moi"}


def _nhan_mac_dinh(k: str) -> str:
    ten = TEN_LOAI.get(k)
    return "Công văn" if k == "cong_van" else (ten.capitalize() if ten else k)


def loai_cua_profile(prof: dict) -> dict[str, dict]:
    kh = prof.get("loai_van_ban")
    if not kh:
        kh = {k: {} for k in TEN_LOAI}
    out: dict[str, dict] = {}
    for k, e in kh.items():
        e = dict(e or {})
        e.setdefault("ten", TEN_LOAI.get(k))
        e.setdefault("nhan", _nhan_mac_dinh(k) if k in TEN_LOAI else (e["ten"] or k).capitalize())
        if "kinh_gui" not in e:
            e["kinh_gui"] = True if k in CO_KINH_GUI else (False if k in KHONG_KINH_GUI else None)
        e.setdefault("muc", [])
        out[k] = e
    return out


def nhan_cua_profile(pid: str) -> dict[str, str]:
    return {k: e["nhan"] for k, e in loai_cua_profile(load_profile(pid)).items()}


@lru_cache(maxsize=1)
def tat_ca_khoa() -> frozenset:
    ks = set(TEN_LOAI)
    for pid in list_profiles("van_ban"):
        ks |= set(loai_cua_profile(load_profile(pid)))
    return frozenset(ks)


def ten_in(prof: dict, khoa: str):
    """Tên loại in trên văn bản (None = không in, như công văn)."""
    e = loai_cua_profile(prof).get(khoa)
    return e["ten"] if e else TEN_LOAI.get(khoa)


def tim_loai(ten_in_hoa: str):
    """Tên loại in hoa (vd. "HỢP ĐỒNG") -> (khóa, id thể thức đầu tiên có loại đó). Không có thì (None, None)."""
    for pid in list_profiles("van_ban"):
        for k, e in loai_cua_profile(load_profile(pid)).items():
            if (e["ten"] or "").upper() == ten_in_hoa and k not in TEN_LOAI:
                return k, pid
    return None, None


@lru_cache(maxsize=1)
def tat_ca_ten() -> dict:
    """Khóa loại -> tên in hoa trên văn bản (tất cả thể thức)."""
    out = {k: v for k, v in TEN_LOAI.items() if v}
    for pid in list_profiles("van_ban"):
        for k, e in loai_cua_profile(load_profile(pid)).items():
            if e.get("ten") and k not in out:
                out[k] = e["ten"]
    return out
