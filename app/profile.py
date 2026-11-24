# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Nạp profile thể thức/slide từ app/profiles/*.yaml.

- Một profile có thể kế thừa profile khác bằng khóa `extends: ten_profile` (chỉ cần ghi phần khác).
- `ap_dung_tuy_chinh(profile, tuy_chinh)` cho người dùng ghi đè từng thông số mà không sửa file yaml:
  chỉ nhận những khóa đã có trong profile gốc, kiểu dữ liệu phải khớp, số phải nằm trong khoảng hợp lý.
"""
from __future__ import annotations

import copy
import math
import re
from functools import lru_cache
from pathlib import Path

import yaml

PROFILE_DIR = Path(__file__).parent / "profiles"
DEFAULT_PROFILE = "nd30"
_RE_TEN = re.compile(r"[a-z0-9_\-]+")
_RE_MAU = re.compile(r"[0-9A-Fa-f]{6}")


def _gop(goc: dict, them: dict) -> dict:
    """Trộn sâu: `them` ghi đè `goc`."""
    out = copy.deepcopy(goc)
    for k, v in them.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _gop(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def _doc_yaml(name: str) -> dict:
    if not _RE_TEN.fullmatch(name):
        raise ValueError(f"Tên profile không hợp lệ: {name!r}")
    path = PROFILE_DIR / f"{name}.yaml"
    if not path.exists():
        raise FileNotFoundError(f"Không tìm thấy profile '{name}'")
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


@lru_cache(maxsize=None)
def _nap(name: str, _chuoi: tuple = ()) -> dict:
    if name in _chuoi:
        raise ValueError(f"Profile kế thừa vòng tròn: {' -> '.join(_chuoi + (name,))}")
    d = _doc_yaml(name)
    cha = d.pop("extends", None)
    if cha:
        d = _gop(_nap(cha, _chuoi + (name,)), d)
    return d


def load_profile(name: str = DEFAULT_PROFILE) -> dict:
    """Trả về bản sao profile đã gộp phần kế thừa (sửa thoải mái, không ảnh hưởng bộ nhớ đệm)."""
    return copy.deepcopy(_nap(name))


def list_profiles(loai: str | None = None) -> list[str]:
    out = []
    for p in sorted(PROFILE_DIR.glob("*.yaml")):
        if p.stem.startswith("_"):
            continue
        if loai is None or _nap(p.stem).get("loai") == loai:
            out.append(p.stem)
    return out


def mo_ta_profile(loai: str | None = None) -> list[dict]:
    """[{id, ten, loai, mo_ta}] cho giao diện chọn thể thức/kiểu slide."""
    out = []
    for pid in list_profiles(loai):
        d = _nap(pid)
        out.append({"id": pid, "ten": d.get("name", pid), "loai": d.get("loai", ""), "mo_ta": d.get("mo_ta", ""),
                    "muc_dich": d.get("muc_dich", "")})
    return out


# --------------------------------------------------------------- ghi đè có kiểm tra
_GIOI_HAN_CHUNG = (-2000.0, 2000.0)
_GIOI_HAN_RIENG = {                       # đuôi khóa -> (min, max)
    "size_pt": (4, 120), "min_size_pt": (4, 120), "sub_size_pt": (4, 120), "title_size_pt": (4, 160),
    "paged_font_pt": (4, 60), "font_sizes_pt": (4, 60),
    "line_spacing": (0.5, 4), "line_exact_pt": (6, 80),
    "max_bullets": (1, 20), "max_lines": (1, 60), "max_chunks": (1, 20),
    "max_lines_per_chunk": (1, 20), "max_words_per_line": (1, 40), "max_words": (1, 200),
    "opacity": (0, 100), "ratio": (0, 1),
}


def _kiem_tra_gia_tri(duong: str, goc, moi):
    """Trả về giá trị hợp lệ (có thể đã ép kiểu) hoặc ném ValueError."""
    if isinstance(goc, bool):
        if not isinstance(moi, bool):
            raise ValueError(f"{duong}: cần giá trị đúng/sai")
        return moi
    if isinstance(goc, (int, float)):
        if isinstance(moi, bool) or not isinstance(moi, (int, float)) or not math.isfinite(moi):
            raise ValueError(f"{duong}: cần là số")
        lo, hi = _GIOI_HAN_RIENG.get(duong.rsplit(".", 1)[-1], _GIOI_HAN_CHUNG)
        if not (lo <= moi <= hi):
            raise ValueError(f"{duong}: nằm ngoài khoảng cho phép [{lo}, {hi}]")
        return type(goc)(moi) if isinstance(goc, int) and float(moi).is_integer() else moi
    if isinstance(goc, str):
        if not isinstance(moi, str):
            raise ValueError(f"{duong}: cần là chuỗi")
        if len(moi) > 400:
            raise ValueError(f"{duong}: quá dài")
        if re.fullmatch(_RE_MAU, goc) and not re.fullmatch(_RE_MAU, moi.lstrip("#")):
            raise ValueError(f"{duong}: cần mã màu 6 ký tự hex, vd 1F3864")
        return moi.lstrip("#") if re.fullmatch(_RE_MAU, goc) else moi
    if isinstance(goc, list):
        if not isinstance(moi, list) or len(moi) > 50:
            raise ValueError(f"{duong}: cần là danh sách")
        if goc and not all(type(x) in (type(goc[0]), int, float) for x in moi if not isinstance(x, bool)):
            raise ValueError(f"{duong}: sai kiểu phần tử")
        return moi
    if goc is None:
        return moi
    raise ValueError(f"{duong}: không cho phép ghi đè")


def ap_dung_tuy_chinh(prof: dict, tuy_chinh: dict | None, bo_qua_la: bool = False) -> dict:
    """Ghi đè thông số profile bằng `tuy_chinh` (dict lồng nhau cùng cấu trúc với yaml).

    Khóa không có trong profile gốc bị từ chối (ValueError), trừ khi bo_qua_la=True (bỏ qua êm).
    Các khóa nhận dạng (loai, name, extends, mo_ta) không được ghi đè.
    """
    if not tuy_chinh:
        return prof
    if not isinstance(tuy_chinh, dict):
        raise ValueError("tuy_chinh phải là đối tượng JSON")
    out = copy.deepcopy(prof)

    def di(goc: dict, moi: dict, duong: str):
        for k, v in moi.items():
            p = f"{duong}.{k}" if duong else str(k)
            if not duong and k in ("loai", "name", "extends", "mo_ta"):
                continue
            if k not in goc:
                if bo_qua_la:
                    continue
                raise ValueError(f"Thông số không tồn tại trong thể thức: {p}")
            if isinstance(goc[k], dict):
                if not isinstance(v, dict):
                    raise ValueError(f"{p}: cần là nhóm thông số")
                di(goc[k], v, p)
            else:
                goc[k] = _kiem_tra_gia_tri(p, goc[k], v)
    di(out, tuy_chinh, "")
    return out
