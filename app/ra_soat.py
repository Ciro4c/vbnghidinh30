# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Rà soát thông tin nhập vào form theo thể thức/quy định đang chọn.

Kết quả là danh sách {muc: loi|canh_bao|goi_y|dat, truong, noi_dung}. Chỉ là kiểm tra theo quy tắc (không dùng AI).
Quy tắc nào áp dụng cho thể thức nào do mục `ra_soat` trong profile quyết định; thành phần thể thức bị bỏ chọn
thì không bị coi là thiếu."""
from __future__ import annotations

import re
from datetime import date

from .loai_vb import loai_cua_profile
from .profile import ap_dung_tuy_chinh, load_profile

# Viết tắt tên loại văn bản (NĐ 30/2020/NĐ-CP, Phụ lục III) dùng đầu phần ký hiệu
VIET_TAT = {
    "quyet_dinh": ["QĐ"], "to_trinh": ["TTr"], "thong_bao": ["TB"], "bao_cao": ["BC"], "ke_hoach": ["KH"],
    "bien_ban": ["BB"], "chi_thi": ["CT"], "huong_dan": ["HD", "HDLN"], "chuong_trinh": ["CTr"], "quy_che": ["QC"],
    "quy_dinh": ["QyĐ", "QĐ"], "nghi_quyet": ["NQ"], "giay_moi": ["GM"], "giay_gioi_thieu": ["GGT"],
    "de_an": ["ĐA"], "phuong_an": ["PA"], "nghi_dinh": ["NĐ"], "thong_tu": ["TT"], "thong_tu_lien_tich": ["TTLT"],
    "cong_dien": ["CĐ"], "thong_cao": ["TC"], "ket_luan": ["KL"],
}
# Loại văn bản không có mục "Kính gửi"
KHONG_KINH_GUI = {"quyet_dinh", "nghi_quyet", "nghi_dinh", "thong_tu", "thong_tu_lien_tich", "quy_che", "quy_dinh",
                  "chi_thi", "thong_cao", "ket_luan", "bien_ban"}
CO_KINH_GUI = {"cong_van", "to_trinh", "bao_cao", "de_an", "phuong_an", "giay_moi"}
RE_QUYEN_HAN = re.compile(r"^(TM\.|KT\.|TL\.|TUQ\.|Q\.)(\s+\S.*)?$")


def _p(muc, truong, nd):
    return {"muc": muc, "truong": truong, "noi_dung": nd}


def ra_soat(f: dict, profile: str, thanh_phan: dict | None = None, tuy_chinh: dict | None = None,
            hom_nay: date | None = None) -> list[dict]:
    try:
        prof = ap_dung_tuy_chinh(load_profile(profile), tuy_chinh or {}, bo_qua_la=True)
    except Exception:
        prof = load_profile(profile)
    rs = prof.get("ra_soat", {}) or {}
    tp = dict(prof.get("thanh_phan", {}))
    tp.update({k: bool(v) for k, v in (thanh_phan or {}).items() if k in tp})
    nhan = prof.get("nhan_thanh_phan", {})
    ds_loai = loai_cua_profile(prof)
    loai = str(f.get("loai_van_ban") or next(iter(ds_loai)))
    out: list[dict] = []
    e_loai = ds_loai.get(loai)
    if e_loai is None:
        out.append(_p("loi", "loai_van_ban", f"Loại “{loai}” không thuộc thể thức này; hãy chọn lại loại văn bản/tài liệu."))
        e_loai = {"muc": [], "kinh_gui": None, "nhan": loai}
    g = lambda k: str(f.get(k) or "").strip()                       # noqa: E731
    on = lambda k: bool(tp.get(k, True))                           # noqa: E731

    # ---- cơ quan
    if on("co_quan"):
        if not g("co_quan_ban_hanh"):
            out.append(_p("loi", "co_quan_ban_hanh", f"Chưa nhập {nhan.get('co_quan', 'tên cơ quan').lower()} (hoặc bỏ chọn thành phần này)."))
        else:
            for k, ten in (("co_quan_ban_hanh", "Cơ quan ban hành"), ("co_quan_chu_quan", "Cơ quan chủ quản")):
                v = g(k)
                if v and v != v.upper() and not rs.get("cho_phep_chu_thuong_co_quan"):
                    out.append(_p("goi_y", k, f"{ten} sẽ được tự đổi thành chữ in hoa."))
            if g("co_quan_chu_quan") and g("co_quan_chu_quan").upper() == g("co_quan_ban_hanh").upper():
                out.append(_p("canh_bao", "co_quan_chu_quan", "Cơ quan chủ quản trùng cơ quan ban hành; chỉ cần ghi một dòng."))

    # ---- số, ký hiệu
    if on("so_ky_hieu"):
        so, kh = g("so"), g("ky_hieu")
        if so and not re.fullmatch(r"\d{1,6}(?:/\d{4})?", so):
            out.append(_p("canh_bao", "so", "Số văn bản nên là số thứ tự (ví dụ 12 hoặc 12/2026), không kèm chữ."))
        if so.isdigit() and int(so) < 10 and rs.get("so_duoi_10_them_0", True):
            out.append(_p("goi_y", "so", f"Số nhỏ hơn 10 sẽ được ghi thành 0{int(so)}."))
        if not kh:
            out.append(_p("goi_y", "ky_hieu", "Chưa nhập ký hiệu; nếu để trống, văn bản chỉ in số."))
        else:
            if kh != kh.upper() and not rs.get("ky_hieu_cho_phep_thuong"):
                out.append(_p("goi_y", "ky_hieu", "Ký hiệu sẽ được tự đổi thành chữ in hoa."))
            if re.search(r"\s", kh):
                out.append(_p("canh_bao", "ky_hieu", "Ký hiệu không nên có dấu cách (dùng gạch nối hoặc gạch chéo)."))
            mau = rs.get("ky_hieu_mau")
            if mau and not re.search(mau, kh.upper()):
                out.append(_p("canh_bao", "ky_hieu", rs.get("ky_hieu_mau_nd") or "Ký hiệu chưa đúng dạng của thể thức đã chọn."))
            if rs.get("ky_hieu_theo_loai", False):
                vt = e_loai.get("viet_tat") or VIET_TAT.get(loai)
                dau = re.split(r"[-/]", kh.upper())[0]
                if vt and not any(dau.startswith(x.upper()) or kh.upper().find("/" + x.upper()) >= 0 or
                                  ("-" + x.upper()) in kh.upper() or ("/" + x.upper()) in kh.upper() for x in vt):
                    out.append(_p("canh_bao", "ky_hieu", f"Văn bản loại này thường có ký hiệu chứa chữ viết tắt {' hoặc '.join(vt)} (ví dụ 01/{vt[0]}-ABC)."))
                if loai == "cong_van" and re.match(r"^(QĐ|TTR|TB|BC|KH|BB|CT|HD|CTR|QC|NQ|GM|ĐA|PA)\b", dau):
                    out.append(_p("goi_y", "ky_hieu", "Công văn thường chỉ ghi chữ viết tắt cơ quan và đơn vị soạn thảo (ví dụ SGDĐT-VP), không có chữ viết tắt loại."))

    # ---- địa danh, ngày
    if on("dia_danh_ngay"):
        if not g("dia_danh"):
            out.append(_p("canh_bao", "dia_danh", "Chưa nhập địa danh (nơi ban hành, ví dụ Cần Thơ); ngày tháng sẽ in không có địa danh."))
        elif re.search(r"[,0-9]|ngày", g("dia_danh"), re.I):
            out.append(_p("canh_bao", "dia_danh", "Địa danh chỉ ghi tên nơi (ví dụ Cần Thơ), không kèm dấu phẩy hay ngày tháng."))
        ng = g("ngay")
        if not ng:
            out.append(_p("loi", "ngay", "Chưa chọn ngày ban hành."))
        else:
            try:
                d = date.fromisoformat(ng[:10])
                ht = hom_nay or date.today()
                if d > ht and not rs.get("cho_phep_ngay_tuong_lai"):
                    out.append(_p("canh_bao", "ngay", "Ngày ban hành ở tương lai; văn bản thường ghi ngày ký, hãy kiểm tra."))
                if d.year < 2000:
                    out.append(_p("canh_bao", "ngay", "Năm ban hành có vẻ sai."))
            except ValueError:
                out.append(_p("loi", "ngay", "Ngày không hợp lệ."))

    # ---- tên loại, trích yếu
    ty = g("trich_yeu")
    if on("trich_yeu") or on("ten_loai"):
        if loai == "cong_van":
            if on("trich_yeu") and not ty:
                out.append(_p("goi_y", "trich_yeu", "Công văn nên có trích yếu “V/v …” ngay dưới số, ký hiệu."))
            if re.match(r"^v/?v\b", ty, re.I) or re.match(r"^về việc", ty, re.I):
                out.append(_p("goi_y", "trich_yeu", "Chỉ nhập nội dung sau “V/v”; hệ thống tự thêm “V/v”."))
        else:
            if on("trich_yeu") and not ty and (e_loai.get("trich_yeu") == "bat_buoc" or (rs.get("trich_yeu_bat_buoc", True) and e_loai.get("trich_yeu") != "khong")):
                vd = e_loai.get("trich_yeu_nhan") or "ví dụ “Về việc ban hành quy chế …”"
                out.append(_p("canh_bao", "trich_yeu", f"{e_loai['nhan']} cần trích yếu/tên gọi: {vd}; in ngay dưới tên loại."))
            if ty and not re.match(r"^(về việc|v/v|về|ban hành|phê duyệt)", ty, re.I) and rs.get("trich_yeu_ve_viec", True) \
                    and loai in ("quyet_dinh", "thong_bao", "to_trinh", "bao_cao", "ke_hoach", "huong_dan", "chi_thi", "ke_hoach_nam_hoc"):
                out.append(_p("goi_y", "trich_yeu", "Trích yếu của loại văn bản này thường mở đầu bằng “Về việc …”."))
        if ty.endswith("."):
            out.append(_p("goi_y", "trich_yeu", "Trích yếu không kết thúc bằng dấu chấm."))
        if len(ty.split()) > 40:
            out.append(_p("canh_bao", "trich_yeu", "Trích yếu quá dài (nên ngắn gọn, dưới 40 từ)."))

    # ---- kính gửi
    kg = f.get("kinh_gui")
    kg = [x for x in (kg if isinstance(kg, list) else str(kg or "").splitlines()) if str(x).strip()]
    if on("kinh_gui"):
        if e_loai.get("kinh_gui") is True and not kg:
            out.append(_p("canh_bao", "kinh_gui", f"{e_loai['nhan']} cần mục “Kính gửi”."))
        if e_loai.get("kinh_gui") is False and kg:
            out.append(_p("canh_bao", "kinh_gui", f"{e_loai['nhan']} không có mục “Kính gửi”; nên bỏ trống."))

    # ---- nội dung
    nd = g("noi_dung")
    if not nd:
        out.append(_p("loi", "noi_dung", "Chưa có nội dung văn bản."))
    elif len(nd) < 20:
        out.append(_p("goi_y", "noi_dung", "Nội dung quá ngắn."))

    elif e_loai.get("muc"):
        toan = "\n".join([nd, ty])
        thieu = []
        for m in e_loai["muc"]:
            try:
                co = re.search(m.get("khop") or ".", toan, re.I)
            except re.error:
                co = True
            if not co:
                thieu.append(m["ten"])
        if thieu:
            out.append(_p("canh_bao", "noi_dung", f"{e_loai['nhan']} thường có các mục: {'; '.join(thieu)} - chưa thấy trong nội dung."))
        if e_loai.get("goi_y") and thieu:
            out.append(_p("goi_y", "noi_dung", e_loai["goi_y"]))

    # ---- đề mục lớn: khung điều hướng + ngắt trang
    if nd:
        from .renderers.cong_van import RE_CHUONG, RE_TIEU_DE_RIENG
        dong = [x.strip() for x in nd.splitlines()]
        co_lon = [x for x in dong if RE_CHUONG.match(x) or RE_TIEU_DE_RIENG.match(x)]
        if co_lon:
            nt = prof.get("ngat_trang") or {}
            if nt.get("chuong") or nt.get("tieu_de"):
                out.append(_p("goi_y", "noi_dung", f"Có {len(co_lon)} đề mục lớn (Chương, Mục lục, Phụ lục...): sẽ hiện trong khung điều hướng và bắt đầu ở trang mới."))
            else:
                out.append(_p("goi_y", "noi_dung", f"Có {len(co_lon)} đề mục lớn (Chương, Mục lục, Phụ lục...): sẽ hiện trong khung điều hướng; thể thức này không tự ngắt trang, bật ở “Tùy chỉnh trình bày” nếu cần."))
        elif re.search(r"(?im)^\s*(mục lục|phụ lục|chương\s+\w+)\b", nd):
            out.append(_p("canh_bao", "noi_dung", "Có dòng giống Mục lục/Phụ lục/Chương nhưng viết chưa đúng dạng (nên viết in hoa riêng một dòng, hoặc “Chương I. Tên”) nên không vào khung điều hướng và không ngắt trang."))

    # ---- chữ ký
    if on("chu_ky"):
        if not g("nguoi_ky"):
            out.append(_p("canh_bao", "nguoi_ky", "Chưa nhập họ tên người ký (để trống thì chừa chỗ ký tay)."))
        if not g("chuc_vu_ky"):
            out.append(_p("canh_bao", "chuc_vu_ky", "Chưa nhập chức vụ người ký."))
        qh = g("quyen_han_ky")
        rx_qh = re.compile(rs["quyen_han_mau"]) if rs.get("quyen_han_mau") else RE_QUYEN_HAN
        if qh and not rx_qh.match(qh.upper()):
            out.append(_p("canh_bao", "quyen_han_ky", "Quyền hạn ký thường bắt đầu bằng TM. (thay mặt), KT. (ký thay), TL. (thừa lệnh), TUQ. (thừa ủy quyền) hoặc Q. (quyền), ví dụ “KT. GIÁM ĐỐC”."))
        if g("nguoi_ky") and re.search(r"\d", g("nguoi_ky")):
            out.append(_p("canh_bao", "nguoi_ky", "Họ tên người ký có chữ số."))
        if qh.upper().startswith(("KT.", "TL.", "TUQ.")) and not g("chuc_vu_ky"):
            out.append(_p("canh_bao", "chuc_vu_ky", "Ký thay/thừa lệnh phải ghi chức vụ của người trực tiếp ký."))

    # ---- nơi nhận
    nn = f.get("noi_nhan")
    nn = [x for x in (nn if isinstance(nn, list) else str(nn or "").splitlines()) if str(x).strip()]
    if on("noi_nhan"):
        if not nn:
            out.append(_p("goi_y", "noi_nhan", "Chưa có nơi nhận."))
        elif rs.get("noi_nhan_luu_vt", True) and not any(re.match(r"^\W*lưu\s*:", x, re.I) for x in nn):
            out.append(_p("goi_y", "noi_nhan", "Dòng cuối nơi nhận thường là “Lưu: VT, …”."))

    # ---- thành phần bị bỏ chọn nhưng vẫn nhập: báo để người dùng biết phần đó không được in
    nhap = {"co_quan": [g("co_quan_ban_hanh")], "so_ky_hieu": [g("so"), g("ky_hieu")], "dia_danh_ngay": [g("dia_danh")],
            "trich_yeu": [ty], "kinh_gui": kg, "chu_ky": [g("nguoi_ky"), g("chuc_vu_ky")], "noi_nhan": nn}
    dau_truong = {"co_quan": "co_quan_ban_hanh", "so_ky_hieu": "so", "dia_danh_ngay": "dia_danh", "trich_yeu": "trich_yeu",
                  "kinh_gui": "kinh_gui", "chu_ky": "nguoi_ky", "noi_nhan": "noi_nhan"}
    for k, gia_tri in nhap.items():
        if not on(k) and any(gia_tri):
            out.append(_p("goi_y", dau_truong[k], f"Thành phần “{nhan.get(k, k)}” đang bỏ chọn nên phần đã nhập sẽ không được in."))

    if not any(x["muc"] in ("loi", "canh_bao") for x in out):
        out.insert(0, _p("dat", "", "Thông tin phù hợp với thể thức đã chọn."))
    return out
