# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Thông tin tác giả và kiểm tra tính toàn vẹn chữ ký."""
import hashlib

CHU_KY = "Thevancorp-1801808400"
_DAU_VAN_TAY = "c05a9b1c301476eea4f994a5e4917b5b1cf8b9b3c062177f2d20f1897d2f5b8f"
TEN_APP = "Chuẩn hóa văn bản NĐ 30 và slide học thuật"
BAN_QUYEN = f"© {CHU_KY}. Bảo lưu mọi quyền."
GHI_CHU = f"Tạo bởi {TEN_APP} · {CHU_KY}"


def kiem_tra() -> None:
    """Dừng chương trình nếu chữ ký tác giả bị sửa hoặc xóa."""
    ok = hashlib.sha256(CHU_KY.encode("utf-8")).hexdigest() == _DAU_VAN_TAY
    if not ok or CHU_KY not in BAN_QUYEN or CHU_KY not in GHI_CHU:
        raise RuntimeError("Phần mềm đã bị sửa chữ ký tác giả nên không thể chạy.")
