# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Tạo file Word mẫu để xem thử, không cần chạy web.
Chạy từ thư mục gốc dự án:  python scripts/tao_mau.py
"""
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.renderers.cong_van import cong_van_to_bytes  # noqa: E402
from app.schemas import CongVan  # noqa: E402

MAU = CongVan(
    co_quan_chu_quan="UBND TỈNH ABC",
    co_quan_ban_hanh="SỞ GIÁO DỤC VÀ ĐÀO TẠO",
    so="5",
    ky_hieu="SGDĐT-VP",
    dia_danh="Cần Thơ",
    ngay=date(2026, 3, 2),
    trich_yeu="Hướng dẫn tổ chức sinh hoạt chuyên môn cấp trường năm học 2026-2027",
    kinh_gui=["Các phòng giáo dục và đào tạo", "Các trường trung học phổ thông trực thuộc"],
    noi_dung=(
        "Thực hiện Kế hoạch số 12/KH-UBND ngày 10 tháng 01 năm 2026 của Ủy ban nhân dân tỉnh về nâng cao chất lượng giáo dục, Sở Giáo dục và Đào tạo hướng dẫn các đơn vị tổ chức sinh hoạt chuyên môn như sau:\n"
        "# 1. Mục đích, yêu cầu\n"
        "- Nâng cao năng lực chuyên môn của đội ngũ giáo viên;\n"
        "- Bảo đảm sinh hoạt chuyên môn thực hiện đúng kế hoạch, thiết thực, hiệu quả.\n"
        "# 2. Nội dung thực hiện\n"
        "## 2.1. Sinh hoạt theo nghiên cứu bài học\n"
        "Các đơn vị lựa chọn nội dung bài học khó, xây dựng kế hoạch dạy học, tổ chức dự giờ và phân tích bài học theo hướng lấy học sinh làm trung tâm.\n"
        "# 3. Tổ chức thực hiện\n"
        "Đề nghị các đơn vị nghiêm túc triển khai thực hiện và báo cáo kết quả về Sở trước ngày 30 tháng 6 năm 2026."
    ),
    quyen_han_ky="KT. GIÁM ĐỐC",
    chuc_vu_ky="PHÓ GIÁM ĐỐC",
    nguoi_ky="Nguyễn Văn A",
    noi_nhan=["Như trên", "Giám đốc (để b/c)", "Lưu: VT, GDTrH"],
)

if __name__ == "__main__":
    out = ROOT / "out"
    out.mkdir(exist_ok=True)
    path = out / "mau-cong-van.docx"
    path.write_bytes(cong_van_to_bytes(MAU))
    print("Đã tạo:", path)
