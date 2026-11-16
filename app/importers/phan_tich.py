# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Nhận diện cấu trúc văn bản hành chính từ danh sách dòng thô (không dùng AI).

Kết quả là các trường để điền vào form; người dùng rà soát lại trước khi xuất.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date

from ..schemas import TEN_LOAI
from .docx_reader import Line

DASH = r"[\-–—•+*]"
RE_SEP = re.compile(r"^[\-_=–—.\s*]+$")
RE_NGAY = re.compile(r"ngày\s*([\d.…]+)\s*tháng\s*([\d.…]+)\s*năm\s*([\d.…]+)", re.I)
RE_SO = re.compile(r"\bSố\s*[:：]\s*(\S*)")
RE_VV = re.compile(r"^(?:V/v|Về việc)\s*[:.]?\s*(.*)$", re.I)
RE_KG = re.compile(r"^Kính\s+gửi\s*[:：]?\s*(.*)$", re.I)
RE_NN = re.compile(r"^Nơi\s+nhận\s*[:：]?\s*$", re.I)
RE_KY = re.compile(r"^(TM|KT|TL|TUQ|Q)\.\s*(.*)$")
RE_PAREN = re.compile(r"^\(.*\)$")
RE_MUC = re.compile(r"^(?:[IVX]+|\d+)\s*[.)]\s+\S")
RE_TIEU_MUC = re.compile(r"^\d+(?:\.\d+)+\.?\s+\S")

RE_PHU_LUC = re.compile(r"^PHỤ\s+LỤC\b", re.I)
RE_DIEU_DONG = re.compile(r"^(Điều\s+\d+[a-zA-Z]?)\s*[:.\-–]\s*(.*)$")
NGAT_TRANG = "[[ngat-trang]]"
DO_DAI_QUOC_HIEU = 60          # dòng quốc hiệu/tiêu ngữ thật rất ngắn; câu dài chỉ nhắc tới tên nước thì không phải

TEN_LOAI_NGUOC = {v: k for k, v in TEN_LOAI.items() if v}


def _bo_dau(s: str) -> str:
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _ti_le_hoa(s: str) -> float:
    chu = [c for c in s if c.isalpha()]
    return sum(c.isupper() for c in chu) / len(chu) if chu else 0.0


def _la_quoc_hieu(text: str) -> bool:
    """Dòng quốc hiệu/tiêu ngữ của đầu văn bản (không nhầm với câu \"Căn cứ ... nước Cộng hòa xã hội chủ nghĩa Việt Nam\")."""
    if len(text) > DO_DAI_QUOC_HIEU:
        return False
    nt = _bo_dau(text)
    return "cong hoa xa hoi chu nghia viet nam" in nt or ("doc lap" in nt and "hanh phuc" in nt)


def _la_hoa(s: str) -> bool:
    return len(s) <= 90 and _ti_le_hoa(s) >= 0.8


def _bo_gach(s: str) -> str:
    return re.sub(rf"^{DASH}\s*", "", s).strip()


def _la_dau_van_ban(parts: list[str]) -> bool:
    """Các phần cách nhau bởi tab có phải là khối đầu văn bản (trái | phải) không."""
    for p in parts:
        nt = _bo_dau(p)
        if ("cong hoa xa hoi" in nt or ("doc lap" in nt and "hanh phuc" in nt)
                or re.match(r"so\s*:", nt) or RE_NGAY.search(p)):
            return True
    return all(_la_hoa(p) for p in parts)


def _tach_ngang(lines: list[Line], gioi_han: int = 25) -> list[Line]:
    """Dòng đầu văn bản dùng tab/khoảng cách lớn để chia trái-phải thì tách ra."""
    out = []
    for i, ln in enumerate(lines):
        parts = [x.strip() for x in re.split(r"\t+", ln.text) if x.strip()]
        if i < gioi_han and len(parts) > 1 and _la_dau_van_ban(parts):
            out.extend(Line(x, ln.bold, ln.italic, ln.align) for x in parts)
        else:
            out.append(ln)
    return out


def _chu_ky(texts: list[str]) -> tuple[str, str, str]:
    ls = [t for t in texts if not RE_PAREN.match(t)]
    quyen = chuc_vu = ten = ""
    if ls and RE_KY.match(ls[0]):
        quyen, ls = ls[0], ls[1:]
    thuong = [i for i, t in enumerate(ls) if not _la_hoa(t)]
    if thuong:
        k = thuong[-1]
        ten = ls[k]
        chuc_vu = " ".join(ls[:k])
    elif len(ls) >= 2:
        chuc_vu, ten = " ".join(ls[:-1]), ls[-1]
    elif ls:
        chuc_vu = ls[0]
    if ten.isupper():
        ten = ten.title()
    return quyen, chuc_vu, ten


def _het_chu_ky(L: list[Line], k: int, gioi_han: int) -> int:
    """Chỉ số dòng ngay sau khối chữ ký bắt đầu ở k. Khối ký ngắn (chức vụ, (ký tên), họ tên);
    gặp bảng, hình, ngắt trang, tiêu đề (Heading) hoặc dòng dài là đã sang phụ lục/nội dung khác."""
    e = k
    while e < gioi_han and e - k < 6:
        ln = L[e]
        if ln.kind != "text" or ln.heading or len(ln.text) > 70 or RE_PHU_LUC.match(ln.text):
            break
        e += 1
    return e


def _doi_ngay(d, m, y) -> str:
    try:
        return date(int(y), int(m), int(d)).isoformat()
    except ValueError:
        return ""


def _noi_dung(body: list[Line], phu_luc: bool = False) -> str:
    """Đổi các dòng thô thành văn bản đánh dấu. Ở phụ lục giữ ngắt trang giữa các biểu mẫu và
    Heading 1/2/3 thành tên phụ lục (giữa) / mục / tiểu mục."""
    out = []
    for ln in body:
        if ln.kind == "ngat":
            if phu_luc and out and out[-1] != NGAT_TRANG:
                out.append(NGAT_TRANG)
            continue                          # ngắt trang trong thân văn bản: để Word tự ngắt
        if ln.kind != "text":                 # dòng bảng/hình: giữ nguyên dạng đánh dấu
            out.append(ln.raw)
            continue
        t = ln.text
        if RE_PAREN.match(t) and len(t) < 40 and not phu_luc:
            continue
        md = RE_DIEU_DONG.match(t)
        if md and not phu_luc and len(t) <= 160:
            out.append(f"{md.group(1)}. {md.group(2)}".strip())   # \"Điều 1: Tên\" -> \"Điều 1. Tên\" (đúng thể thức)
            continue
        if phu_luc and ln.heading:
            out.append({1: "> ", 2: "# "}.get(ln.heading, "## ") + t)
            continue
        gach = re.match(rf"^{DASH}\s+(.*)$", t)
        ngan_khong_cham = len(t) <= 100 and not t.endswith((".", ";", ",", ":"))
        la_muc = (ln.bold or _la_hoa(t) or ngan_khong_cham)
        if (ln.align == "center" and ln.bold and len(t) <= 100
                and (_la_hoa(t) or t.endswith(":"))):
            out.append("> " + t)
        elif ln.heading:
            out.append(("# " if ln.heading <= 1 else "## ") + t)
        elif gach:
            out.append("- " + gach.group(1))
        elif ln.bullet:
            out.append("- " + t)
        elif RE_TIEU_MUC.match(t) and la_muc and len(t) <= 100:
            out.append("## " + t)
        elif RE_MUC.match(t) and la_muc and len(t) <= 100:
            out.append("# " + t)
        elif ln.bold and ngan_khong_cham and (phu_luc or not _la_hoa(t)) and ln.align != "center":
            out.append("# " + t)
        else:
            out.append(t)
    return "\n".join(out)


def phan_tich(lines: list[Line]) -> tuple[dict, list[str]]:
    warns: list[str] = []
    L = _tach_ngang([l for l in lines if l.text.strip()])
    for l in L:                       # tab còn lại ngoài phần đầu chỉ là thụt lề
        l.text = l.text.replace("\t", " ").strip()
    n = len(L)
    H = min(n, 25)
    used: set[int] = set()
    f = {"loai_van_ban": "cong_van", "co_quan_chu_quan": "", "co_quan_ban_hanh": "",
         "so": "", "ky_hieu": "", "dia_danh": "", "ngay": "", "trich_yeu": "",
         "kinh_gui": [], "noi_dung": "", "quyen_han_ky": "", "chuc_vu_ky": "",
         "nguoi_ky": "", "noi_nhan": [], "phu_luc": "", "phu_luc_ngang": False}

    # --- quốc hiệu, tiêu ngữ, đường kẻ: bỏ qua (hệ thống tự dựng lại)
    for i in range(H):
        if _la_quoc_hieu(L[i].text) or RE_SEP.match(L[i].text):
            used.add(i)

    # --- số, ký hiệu, địa danh, ngày tháng
    i_so = i_ngay = None
    for i in range(H):
        t = L[i].text
        if t.startswith("Căn cứ"):          # \"... ngày 17 tháng 6 năm 2020\" trong căn cứ không phải ngày ban hành
            continue
        m = RE_SO.search(t)
        if m and i_so is None:
            tok = m.group(1)
            so, _, ky = tok.partition("/")
            f["so"] = so if re.fullmatch(r"\d+", so) else ""
            f["ky_hieu"] = ky.strip(" .…")
            i_so = i
            used.add(i)
        m = RE_NGAY.search(t)
        if m and i_ngay is None:
            f["ngay"] = _doi_ngay(*m.groups())
            f["dia_danh"] = re.split(r"\t|\s{3,}", t[:m.start()])[-1].strip(" ,")
            if not f["ngay"]:
                warns.append("Ngày tháng trong file chưa đầy đủ, hãy chọn ngày ban hành.")
            i_ngay = i
            used.add(i)
    if i_so is None:
        warns.append("Không tìm thấy dòng “Số: …/…”, hãy điền số và ký hiệu.")
    if i_ngay is None:
        warns.append("Không tìm thấy dòng ngày tháng, hãy chọn ngày ban hành.")

    # --- tên loại + trích yếu
    i_loai = None
    for i in range(min(n, 40)):
        key = TEN_LOAI_NGUOC.get(L[i].text.strip().upper().rstrip(":"))
        if key and i not in used and L[i].text.isupper() and not L[i].text.endswith(":"):
            f["loai_van_ban"], i_loai = key, i
            used.add(i)
            break
    if i_loai is not None:
        j = i_loai + 1
        parts = []
        while j < n and len(parts) < 2 and j not in used:
            t = L[j].text
            if RE_KG.match(t) or t.startswith("Căn cứ") or RE_MUC.match(t):
                break
            if _la_hoa(t) and not RE_VV.match(t):
                break
            if parts and not (L[j].align == "center" and L[j - 1].align == "center"):
                break
            if not parts and not (RE_VV.match(t) or L[j].bold or L[j].align == "center"):
                break
            parts.append(t)
            used.add(j)
            j += 1
        f["trich_yeu"] = " ".join(parts)
    else:
        for i in range(min(n, 40)):
            m = RE_VV.match(L[i].text)
            if m and i not in used:
                f["trich_yeu"] = m.group(1).strip()
                used.add(i)
                break

    # --- kính gửi
    for i in range(min(n, 45)):
        m = RE_KG.match(L[i].text)
        if m and i not in used:
            used.add(i)
            items = [m.group(1)] if m.group(1).strip() else []
            j = i + 1
            while j < n and (re.match(rf"^{DASH}\s*", L[j].text) or
                             (items and items[-1].rstrip().endswith(";") and j - i < 12)):
                items.append(L[j].text)
                used.add(j)
                j += 1
            f["kinh_gui"] = [_bo_gach(x).rstrip(";.,") for x in items if _bo_gach(x)]
            break

    # --- cơ quan ban hành (các dòng chữ in hoa đứng trước số/ngày)
    gioi_han_cq = min([x for x in (i_so, i_ngay, i_loai) if x is not None] or [min(n, 6)])
    cq = [i for i in range(gioi_han_cq) if i not in used and _la_hoa(L[i].text)]
    for i in cq:
        used.add(i)
    if len(cq) == 1:
        f["co_quan_ban_hanh"] = L[cq[0]].text.upper()
    elif len(cq) >= 2:
        f["co_quan_chu_quan"] = L[cq[0]].text.upper()
        f["co_quan_ban_hanh"] = "\n".join(L[i].text.upper() for i in cq[1:])
        warns.append("Đã đoán dòng đầu là cơ quan chủ quản, các dòng sau là cơ quan ban hành; "
                     "hãy kiểm tra.")
    else:
        warns.append("Không nhận diện được cơ quan ban hành, hãy điền thủ công.")

    bat_dau = (max(used) + 1) if used else 0

    # --- nơi nhận + chữ ký + phụ lục (phần nằm sau chữ ký)
    ket_thuc = n
    bat_dau_pl = n                                   # chỉ số dòng đầu của phụ lục (n = không có)
    i_nn = next((i for i in range(bat_dau, n) if RE_NN.match(L[i].text)), None)   # \"Nơi nhận\" đầu tiên
    if i_nn is not None:
        k = i_nn + 1
        while k < n and not (RE_KY.match(L[k].text) or
                             (_la_hoa(L[k].text) and not re.match(rf"^{DASH}", L[k].text))):
            if L[k].kind != "text" or L[k].heading:   # gặp bảng/tiêu đề: đã sang phụ lục, không có khối ký
                break
            k += 1
        e = _het_chu_ky(L, k, n)
        f["noi_nhan"] = [_bo_gach(L[x].text).rstrip(";.,") for x in range(i_nn + 1, k)
                         if L[x].kind == "text" and _bo_gach(L[x].text)]
        f["quyen_han_ky"], f["chuc_vu_ky"], f["nguoi_ky"] = _chu_ky(
            [L[x].text for x in range(k, e) if L[x].kind == "text"])
        ket_thuc, bat_dau_pl = i_nn, e
    else:
        warns.append("Không thấy mục “Nơi nhận”, hãy điền nơi nhận.")
        i_pl = next((i for i in range(bat_dau, n) if L[i].kind == "text" and RE_PHU_LUC.match(L[i].text)), n)
        h = i_pl                                     # chỉ tìm chữ ký trước phụ lục
        k = next((i for i in range(max(bat_dau, h - 6), h)
                  if L[i].kind == "text" and (RE_KY.match(L[i].text) or _la_hoa(L[i].text))), None)
        if k is not None:
            e = _het_chu_ky(L, k, h)
            f["quyen_han_ky"], f["chuc_vu_ky"], f["nguoi_ky"] = _chu_ky(
                [L[x].text for x in range(k, e) if L[x].kind == "text"])
            ket_thuc = k
        else:
            ket_thuc = h
        bat_dau_pl = i_pl
    if not f["nguoi_ky"]:
        warns.append("Chưa nhận diện được họ tên người ký, hãy kiểm tra phần chữ ký.")

    # --- phụ lục: mọi thứ sau chữ ký được giữ nguyên, không đưa vào khối chữ ký
    pl = L[bat_dau_pl:]
    while pl and pl[0].kind == "ngat":
        pl = pl[1:]
    while pl and pl[-1].kind == "ngat":
        pl = pl[:-1]
    if any(x.kind != "ngat" for x in pl):
        f["phu_luc"] = _noi_dung(pl, phu_luc=True)
        f["phu_luc_ngang"] = any(x.ngang for x in pl)
        n_bang = sum(1 for x in pl if x.kind == "bang")
        warns.append(f"Có phần phụ lục sau chữ ký ({len(pl)} dòng, {n_bang} bảng): đã giữ nguyên ở trang riêng"
                     f"{', khổ giấy ngang' if f['phu_luc_ngang'] else ''}. Hãy xem lại phần Phụ lục trong form.")
    f["noi_dung"] = _noi_dung(L[bat_dau:ket_thuc])
    ngoai = [i for i, ln in enumerate(L) if ln.kind not in ("text", "ngat")
             and not (bat_dau <= i < ket_thuc) and i < bat_dau_pl]
    if ngoai:
        warns.append(f"Có {len(ngoai)} hình/bảng ở phần đầu hoặc phần chữ ký, nơi nhận (logo, con dấu, "
                     "chữ ký...) không được chuyển; nếu cần hãy chèn lại thủ công trong Word.")
    if not f["noi_dung"]:
        warns.append("Không tách được phần nội dung.")
    if f["loai_van_ban"] != "cong_van" and f["kinh_gui"]:
        warns.append("Văn bản này (quyết định, thông báo...) theo NĐ 30 không có dòng “Kính gửi”; đã giữ lại "
                     "theo file gốc, hãy xóa ô Kính gửi nếu không cần.")
    if f["loai_van_ban"] == "cong_van" and not f["kinh_gui"]:
        warns.append("Không thấy dòng “Kính gửi”; nếu đây không phải công văn, hãy chọn đúng loại văn bản.")
    return f, warns
