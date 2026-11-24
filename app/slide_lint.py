# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Kiểm tra slide và thông số trình bày theo chuẩn (ngưỡng lấy từ profile `standard`, có thể chỉnh).
Trả về danh sách phát hiện {muc: loi|canh_bao|goi_y|dat, ma, slide, noi_dung}; không sửa gì."""
from __future__ import annotations

import re

RE_DONG_TU = re.compile(
    r"\b(là|có|tăng|giảm|đạt|cần|giúp|làm|cho|tạo|đưa|mang|vượt|chiếm|gồm|bao gồm|dẫn|giữ|tiết kiệm|"
    r"đã|sẽ|đang|phải|nên|vì|khiến|cải thiện|rút|nâng|mở rộng|duy trì|ảnh hưởng|quyết định)\b", re.I)
RE_CHI_DE_MUC = re.compile(r"^(giới thiệu|tổng quan|nội dung|kết luận|mở đầu|đặt vấn đề|kết quả|phương pháp|"
                           r"tài liệu tham khảo|cảm ơn|hỏi đáp|q&a|agenda|overview|introduction)\b", re.I)


def _lin(c: int) -> float:
    c /= 255
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def do_sang(hexc: str) -> float:
    h = hexc.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def ti_le_tuong_phan(a: str, b: str) -> float:
    la, lb = sorted((do_sang(a), do_sang(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def xam(hexc: str) -> str:
    """Mô phỏng ảnh đơn sắc (thang xám theo độ sáng cảm nhận)."""
    h = hexc.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    y = int(round(0.299 * r + 0.587 * g + 0.114 * b))
    return f"{y:02X}{y:02X}{y:02X}"


def _tu(s: str) -> int:
    return len(s.split())


def kiem_tra_mau(prof: dict) -> list[dict]:
    out = []
    c = prof.get("colors", {})
    st = prof.get("standard", {}).get("contrast", {})
    muc = st.get("level", "AA")
    ngu_body = 7.0 if muc == "AAA" else st.get("body", 4.5)
    ngu_lon = 4.5 if muc == "AAA" else st.get("large", 3.0)
    cap = [("text", "bg", "Chữ nội dung trên nền", ngu_body),
           ("primary", "bg", "Tiêu đề trên nền", ngu_lon),
           ("muted", "bg", "Chữ chân trang/nguồn trên nền", ngu_body),
           ("on_primary", "primary", "Chữ trên nền tối (slide mở đầu/kết)", ngu_body),
           ("highlight", "bg", "Màu nhấn từ khóa trên nền", ngu_lon),
           ("accent", "bg", "Màu đồ họa (đường, khối) trên nền", st.get("graphics", 3.0))]
    for a, b, ten, ngu in cap:
        if a in c and b in c:
            try:
                r = ti_le_tuong_phan(c[a], c[b])
            except Exception:
                continue
            ok = r >= ngu
            out.append({"muc": "dat" if ok else "canh_bao", "ma": "tuong_phan",
                        "slide": 0, "noi_dung": f"{ten}: {r:.2f}:1 (cần ≥ {ngu})"
                        + ("" if ok else " – chưa đạt, hãy dùng màu đậm hơn hoặc chỉ dùng cho chữ lớn/đồ họa")})
            if st.get("grayscale", True):
                rg = ti_le_tuong_phan(xam(c[a]), xam(c[b]))
                if rg < ngu and ok:
                    out.append({"muc": "canh_bao", "ma": "don_sac", "slide": 0,
                                "noi_dung": f"{ten}: khi in đen trắng chỉ còn {rg:.2f}:1 (cần ≥ {ngu})"})
    return out


def kiem_tra_thong_so(prof: dict) -> list[dict]:
    out = []
    st = prof.get("standard", {})
    ty = st.get("typography", {})
    zo = st.get("zones", {})
    ly = prof.get("layout", {})

    def khoang(ten, v, lo, hi):
        if lo is not None and hi is not None and not (lo <= v <= hi):
            out.append({"muc": "goi_y", "ma": "thong_so", "slide": 0,
                        "noi_dung": f"{ten} = {v} nằm ngoài khoảng tham khảo {lo}–{hi} (vẫn dùng được)"})
    if ty:
        khoang("Cỡ chữ tiêu đề (pt)", prof["title"]["size_pt"], ty.get("title_min_pt"), ty.get("title_max_pt"))
        khoang("Cỡ chữ nội dung (pt)", prof["body"]["size_pt"], ty.get("body_min_pt"), ty.get("body_max_pt"))
        khoang("Cỡ chữ chân trang (pt)", prof["footer"]["size_pt"], ty.get("footer_min_pt"), ty.get("footer_max_pt"))
    if ly.get("use_percent") and zo:
        khoang("Vùng tiêu đề (%)", ly["title_zone_pct"], zo.get("title_zone_min_pct"), zo.get("title_zone_max_pct"))
        khoang("Vùng chân trang (%)", ly["footer_zone_pct"], zo.get("footer_zone_min_pct"), zo.get("footer_zone_max_pct"))
        khoang("Lề hai bên (%)", ly["side_margin_pct"], zo.get("side_margin_min_pct"), zo.get("side_margin_max_pct"))
        tong = ly["title_zone_pct"] + ly["center_zone_pct"] + ly["footer_zone_pct"]
        if abs(tong - 100) > 1:
            out.append({"muc": "canh_bao", "ma": "thong_so", "slide": 0,
                        "noi_dung": f"Ba vùng tiêu đề/trọng tâm/chân trang cộng {tong}% (nên bằng 100%)"})
    return out


def kiem_tra_slide(slides, prof: dict) -> list[dict]:
    out = []
    st = prof.get("standard", {})
    at, ck = st.get("action_title", {}), st.get("chunking", {})
    for i, sl in enumerate(slides, 1):
        t = (sl.title or "").strip()
        if at.get("enabled", True) and t:
            n = _tu(t)
            if RE_CHI_DE_MUC.match(t) and n <= 4:
                out.append({"muc": "goi_y", "ma": "action_title", "slide": i,
                            "noi_dung": f"“{t}” là nhãn chủ đề; nên viết thành câu kết luận (ví dụ nêu kết quả hoặc đề xuất)"})
            elif n < at.get("min_words", 6) or not RE_DONG_TU.search(t):
                out.append({"muc": "goi_y", "ma": "action_title", "slide": i,
                            "noi_dung": f"“{t}” có thể chưa là câu hoàn chỉnh ({n} từ); Action Title nên nêu rõ kết luận"})
            elif n > at.get("max_words", 22):
                out.append({"muc": "goi_y", "ma": "action_title", "slide": i,
                            "noi_dung": f"Tiêu đề dài {n} từ, nên ≤ {at.get('max_words', 22)}"})
        if not sl.items:
            continue
        mx = ck.get("max_chunks", 5)
        if len(sl.items) > mx:
            out.append({"muc": "canh_bao", "ma": "chunking", "slide": i,
                        "noi_dung": f"{len(sl.items)} ý trên một slide (tối đa {mx}); sẽ được tách sang slide tiếp"})
        dai = [it for it in sl.items if _tu(it[1]) > ck.get("max_words_per_line", 7) * ck.get("max_lines_per_chunk", 5)]
        for it in dai[:2]:
            out.append({"muc": "canh_bao", "ma": "chunking", "slide": i,
                        "noi_dung": f"Ý dài {_tu(it[1])} từ: “{it[1][:50]}…” — nên rút thành từ khóa, phần còn lại để ghi chú"})
        if not any(a for a in (sl.images, sl.tables)) and len(sl.items) >= 3:
            ngan = [it for it in sl.items if it[0] == 0 and _tu(it[1]) <= ck.get("max_words_per_line", 7)]
            if len(ngan) == len(sl.items):
                out.append({"muc": "goi_y", "ma": "bo_cuc", "slide": i,
                            "noi_dung": "Các ý ngắn: có thể dùng bố cục thẻ/dòng thời gian thay cho gạch đầu dòng"})
    return out


def phan_tich(slides, prof: dict) -> dict:
    ds = kiem_tra_mau(prof) + kiem_tra_thong_so(prof) + kiem_tra_slide(slides, prof)
    dem = {"canh_bao": 0, "goi_y": 0, "dat": 0, "loi": 0}
    for d in ds:
        dem[d["muc"]] = dem.get(d["muc"], 0) + 1
    return {"phat_hien": ds, "tong": dem, "so_slide": len(slides),
            "tuan_thu": dem["canh_bao"] == 0 and dem["loi"] == 0}
