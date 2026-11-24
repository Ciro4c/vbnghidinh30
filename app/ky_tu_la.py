# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Dò (và dọn) ký tự, dấu vết lạ thường gặp khi dán văn bản từ chat AI hoặc từ file PDF/Word hỏng.

Đây là dò *dấu vết kỹ thuật* (ký tự vô hình, markdown còn sót, mã trích dẫn, emoji, chữ giống chữ Latin
nhưng thuộc bảng chữ khác, câu dẫn của chatbot, chỗ giữ chỗ chưa điền...). Không phải công cụ kết luận
"văn bản do AI viết": không có cách nào làm việc đó đáng tin cậy, nên module này không đưa ra kết luận như vậy.

Dùng:
    ds = quet(text)                       # danh sách PhatHien
    sach, thong_ke = lam_sach(text)       # văn bản đã dọn + số chỗ đã sửa theo từng loại
Mọi thao tác dọn đều có thể bật/tắt qua `tuy_chon` (xem MAC_DINH).
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass

# ------------------------------------------------------------------ danh mục loại
# mã -> (mức, mô tả, tự sửa được)
LOAI: dict[str, tuple[str, str, bool]] = {
    "vo_hinh": ("cao", "Ký tự vô hình (khoảng trắng rộng 0, điều khiển chiều chữ, ký tự thẻ ẩn...)", True),
    "khoang_trang_la": ("trung_binh", "Khoảng trắng đặc biệt (NBSP, khoảng trắng hẹp, khoảng trắng CJK...)", True),
    "dieu_khien": ("cao", "Ký tự điều khiển", True),
    "thay_the": ("cao", "Ký tự thay thế �: file bị lỗi mã hóa, chữ gốc đã mất", False),
    "loi_ma_hoa": ("cao", "Chữ bị lỗi mã hóa (mojibake, vd “Ã¡”, “áº¡”)", True),
    "unicode_to_hop": ("trung_binh", "Dấu tiếng Việt kiểu tổ hợp (NFD), dễ lỗi khi tìm kiếm/in", True),
    "chu_la_giong_latin": ("cao", "Chữ Cyrillic/Hy Lạp giống chữ Latin lẫn trong từ", True),
    "chu_toan_rong": ("trung_binh", "Chữ toàn chiều rộng hoặc chữ in kiểu toán học (Ａ, 𝗔)", True),
    "ligature": ("thap", "Chữ ghép (ﬁ, ﬂ...) thường sinh ra khi sao chép từ PDF", True),
    "markdown": ("trung_binh", "Ký hiệu markdown còn sót (**đậm**, ###, `mã`, ---)", True),
    "trich_dan_ai": ("cao", "Mã trích dẫn/nguồn của chatbot (oaicite, turn0search0, 【…†source】, [cite…])", True),
    "url_theo_doi": ("thap", "Đường dẫn có tham số theo dõi của chatbot (utm_source=chatgpt.com...)", True),
    "emoji": ("trung_binh", "Emoji, biểu tượng trang trí", True),
    "ky_tu_vung_rieng": ("thap", "Ký tự vùng riêng (thường là bullet của font Symbol/Wingdings)", False),
    "cau_dan_chatbot": ("trung_binh", "Câu dẫn/kết của chatbot (“Dưới đây là…”, “Hy vọng điều này hữu ích…”)", False),
    "giu_cho": ("trung_binh", "Chỗ giữ chỗ chưa điền ([Tên cơ quan], {{…}}, XXX...)", False),
    "cho_ky_dau": ("thap", "Chỗ giữ chỗ chữ ký/con dấu, có thể lồng ảnh chữ ký/mộc vào", False),
    "gach_dai": ("thap", "Dấu gạch dài — (kiểu chatbot hay dùng), văn bản hành chính thường dùng “-” hoặc dấu phẩy", True),
    "ngoac_kep_lan": ("thap", "Dấu ngoặc kép thẳng \" lẫn dấu cong “ ”", False),
    "khoang_trang_thua": ("thap", "Khoảng trắng thừa (nhiều dấu cách liền nhau, cuối dòng)", True),
    "dau_cau_sai": ("thap", "Dấu câu lặp hoặc cách sai (,, ..  ,)", True),
    "so_kieu_my": ("thap", "Số dùng dấu phẩy ngăn nghìn kiểu Mỹ (1,000); văn bản Việt Nam dùng dấu chấm", False),
}

MAC_DINH: dict[str, bool] = {      # tuỳ chọn dọn: bật/tắt từng nhóm
    "xoa_vo_hinh": True,
    "doi_khoang_trang": True,
    "xoa_dieu_khien": True,
    "sua_loi_ma_hoa": True,
    "chuan_unicode_nfc": True,
    "sua_chu_la": True,
    "chuan_chu_toan_rong": True,
    "go_ligature": True,
    "go_markdown": True,
    "go_trich_dan_ai": True,
    "go_url_theo_doi": True,
    "go_emoji": True,
    "doi_gach_dai": False,        # chỉ báo, không tự đổi (có thể là dấu gạch có chủ ý)
    "go_cau_chatbot": False,      # xóa cả dòng: chỉ làm khi người dùng bật
    "don_khoang_trang": True,
    "sua_dau_cau": True,
}

# ------------------------------------------------------------------ bảng ký tự
_VO_HINH = re.compile("[­͏؜᠎​-‏‪-‮⁠-⁯﻿"
                      "︀-︎\U000e0000-\U000e007f]")
_KHOANG_TRANG = re.compile("[   -   　]")
_DIEU_KHIEN = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
_VUNG_RIENG = re.compile("[-]")
_TO_HOP = re.compile("[̀-ͯ]")
_TOAN_RONG = re.compile("[！-～\U0001d400-\U0001d7ff]")
_LIGATURE = re.compile("[ﬀ-ﬆ]")
_EMOJI_CHU = re.compile("[⌀-⏿☀-➿⬀-⯿←-⇿\U0001f000-\U0001faff️]")
_EMOJI_GIU = set("☐☑☒←→↔•")        # ☐ ☑ ☒ ← → ↔ giữ lại nếu dùng làm ô chọn/mũi tên
_GACH_DAI = re.compile(r"\s*[—―]\s*")
_NGOAC_CONG = re.compile("[“”]")
_NGOAC_THANG = re.compile('"')
_SO_MY = re.compile(r"(?<![\d.,])\d{1,3}(?:,\d{3})+(?![\d,]|\.\d{3})")

# chữ giống Latin -> Latin
_GIONG = {
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i", "ј": "j", "ѕ": "s", "ԁ": "d",
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T",
    "Х": "X", "І": "I", "Ј": "J", "Ѕ": "S", "Ү": "Y",
    "ο": "o", "Ο": "O", "Α": "A", "Β": "B", "Ε": "E", "Ζ": "Z", "Η": "H", "Ι": "I", "Κ": "K", "Μ": "M",
    "Ν": "N", "Τ": "T", "Υ": "Y", "Χ": "X", "ν": "v",
}
_DICH_GIONG = str.maketrans(_GIONG)
_CHU_GIONG = re.compile("[" + "".join(_GIONG) + "]")
_TU = re.compile(r"\w+", re.UNICODE)
_LA_LATIN = re.compile(r"[A-Za-zÀ-ỹ]")

_LOI_MA_HOA = re.compile("(?:[ÃÂ][\u0080-¿]|Ä[\u0080-¿]|Æ[\u0080-¿]|áº|á»|â€[\u0080-¿™œ“”˜])")

# trích dẫn do chatbot chèn
_TRICH_DAN = [
    re.compile(r"[^]*"),                                   # citeturn0search0
    re.compile(r":?contentReference\[oaicite:\d+\]\{index=\d+\}"),
    re.compile(r"\[oaicite:\d+\](?:\{index=\d+\})?"),
    re.compile(r"\b(?:turn\d+(?:search|view|news|image|file|forecast|product)\d+)\b"),
    re.compile(r"【[^】\n]{0,80}†[^】\n]{0,80}】"),
    re.compile(r"\[cite(?:_start|_end)?(?::\s*[\d,\s\-]+)?\]"),
    re.compile(r"\[(?:source|nguồn)\s*:\s*\d+(?:\s*,\s*\d+)*\]", re.I),
    re.compile(r"\bcite_?(?:start|end)\b"),
]
_URL_THEO_DOI = re.compile(r"[?&]utm_source=(?:chatgpt\.com|openai|claude\.ai|gemini|perplexity)[^\s&)]*", re.I)

# câu dẫn/kết của chatbot (so khớp theo từng dòng)
_CAU_CHATBOT = [re.compile(p, re.I) for p in (
    r"^\s*(?:chắc chắn rồi|tất nhiên rồi?|dĩ nhiên rồi?|tuyệt vời|được rồi|chắc chắn)\s*[!,.:]",
    r"^\s*(?:dưới đây|sau đây) là (?:một |các )?(?:bản |mẫu |phiên bản |nội dung |gợi ý |dự thảo )[^\n]{0,100}[:：]\s*$",
    r"hy vọng (?:điều|bản|mẫu|nội dung)[^\n]{0,50}(?:hữu ích|giúp ích)",
    r"nếu (?:bạn|anh|chị)[^\n]{0,40}(?:cần|muốn)[^\n]{0,80}(?:hãy|vui lòng) (?:cho (?:tôi|mình)|nói với (?:tôi|mình))",
    r"(?:bạn|anh|chị) có muốn (?:tôi|mình)",
    r"là một (?:mô hình|trợ lý)(?: ngôn ngữ| AI| ảo)?",
    r"\bas an ai\b|\bi hope this helps\b|\bhere(?:'s| is) (?:a|an|the|your)\b|^\s*(?:certainly|sure)[!,.]",
    r"(?:bản|văn bản) (?:dự thảo )?(?:trên|này) chỉ mang tính (?:tham khảo|minh họa)",
    r"(?:vui lòng|hãy|bạn nên) (?:kiểm tra|điều chỉnh|chỉnh sửa) (?:lại )?(?:thông tin|nội dung|số liệu)[^\n]{0,40}(?:phù hợp|thực tế)",
)]

_GIU_CHO = [re.compile(p, re.I) for p in (
    r"\[[^\]\n]{0,50}(?:tên|số|ngày|địa chỉ|địa danh|cơ quan|chức vụ|họ|nội dung|điền|nhập|trích yếu|người ký|đơn vị|"
    r"insert|name|date|address)[^\]\n]{0,50}\]",
    r"\{\{[^}\n]{1,60}\}\}",
    r"<<[^>\n]{1,60}>>",
    r"(?<![A-Za-z0-9])X{3,}(?![A-Za-z0-9])",
)]
_CHO_KY_DAU = re.compile(
    r"[(\[]\s*(?:đã ký(?: số)?|ký tên(?:,?\s*đóng dấu)?|ký,?\s*(?:ghi rõ họ tên|đóng dấu)|chữ ký(?: số)?|"
    r"con dấu|đóng dấu|dấu đỏ|đã đóng dấu)[^)\]\n]{0,25}[)\]]", re.I)

_MD_DAM = re.compile(r"(\*\*|__)(?=\S)(.+?)(?<=\S)\1")
_MD_NGHIENG = re.compile(r"(?<![\w*])\*(?=[^\s*])([^*\n]+?)(?<=[^\s*])\*(?![\w*])")
_MD_CODE = re.compile(r"`+([^`\n]+)`+")
_MD_TIEUDE = re.compile(r"^\s{0,3}(#{1,6})\s+(.*\S)\s*$")
_MD_TIEUDE_3 = re.compile(r"^\s{0,3}(#{3,6})\s+(.*\S)\s*$")
_MD_NGANG = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,}|={3,})\s*$")
_MD_GACH_SAO = re.compile(r"^(\s*)[*+]\s+(?=\S)")
_MD_BANG_PHAN_CACH = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
_MD_LIEN_KET = re.compile(r"\[([^\]\n]+)\]\((?:https?://|mailto:)[^)\s]+\)")

_DAU_CAU_LAP = re.compile(r"(?<![.\d…])(?:,{2,}|;{2,}|:{2,}|\.{2}(?!\.))")
_CACH_TRUOC_DAU = re.compile(r"[ \t]+([,;:.!?)])")
_THIEU_CACH_SAU = re.compile(r"([,;])(?=[^\s\d\)\]”\"'’.,;])")


@dataclass
class PhatHien:
    ma: str
    muc: str
    dong: int            # 1 = dòng đầu (0 nếu không xác định)
    cot: int
    mau: str             # đoạn trích, ký tự vô hình viết thành ‹U+200B›
    mo_ta: str
    tu_sua: bool
    trang: int = 0       # số trang (PDF), 0 nếu không phải PDF

    def to_dict(self) -> dict:
        return asdict(self)


def _hien(s: str) -> str:
    """Hiện ký tự vô hình/lạ dưới dạng ‹U+XXXX› để đọc được."""
    out = []
    for ch in s:
        cat = unicodedata.category(ch)
        if cat in ("Cf", "Cc", "Co", "Cs", "Zs", "Zl", "Zp") and ch != " ":
            out.append(f"‹U+{ord(ch):04X}›")
        else:
            out.append(ch)
    return "".join(out)


def _mau(dong: str, a: int, b: int, bien: int = 18) -> str:
    return _hien(dong[max(0, a - bien):a]) + "⟦" + _hien(dong[a:b]) + "⟧" + _hien(dong[b:b + bien])


def _tuy_chon(tc: dict | None) -> dict:
    out = dict(MAC_DINH)
    for k, v in (tc or {}).items():
        if k in out:
            out[k] = bool(v)
    return out


def _la_emoji(ch: str) -> bool:
    return bool(_EMOJI_CHU.match(ch)) and ch not in _EMOJI_GIU


# ------------------------------------------------------------------ dò
def quet(text: str, che_do: str = "goc", trang: int = 0, toi_da: int = 300,
         dong_dau: int = 1) -> list[PhatHien]:
    """Dò một đoạn văn bản.

    che_do="goc": văn bản thô (từ Word, TXT, PDF): ký hiệu markdown như `#`, `-`, `>` đều được coi là lạ.
    che_do="form": nội dung trong ô soạn thảo của ứng dụng, nơi `# `, `## `, `- `, `> `, `| |` là ký hiệu riêng của app.
    """
    ds: list[PhatHien] = []

    da_lay: dict[tuple[str, int], list[tuple[int, int]]] = {}

    def them(ma, dong, a, b, ln):
        if len(ds) >= toi_da:
            return
        cac = da_lay.setdefault((ma, dong), [])
        if any(a < y + 3 and b > x - 3 for x, y in cac):        # trùng/liền sát chỗ đã báo cùng loại: gộp
            return
        cac.append((a, b))
        muc, mo_ta, tu_sua = LOAI[ma]
        ds.append(PhatHien(ma, muc, dong, a + 1, _mau(ln, a, b), mo_ta, tu_sua, trang))

    dong_goc = text.split("\n")
    n_ngoac_cong = n_ngoac_thang = 0
    for i, ln in enumerate(dong_goc):
        so = dong_dau + i
        ln = ln.rstrip("\r")
        if not ln.strip():
            continue
        for m in _VO_HINH.finditer(ln):
            them("vo_hinh", so, m.start(), m.end(), ln)
        for m in _KHOANG_TRANG.finditer(ln):
            them("khoang_trang_la", so, m.start(), m.end(), ln)
        for m in _DIEU_KHIEN.finditer(ln.replace("\t", "")):
            them("dieu_khien", so, m.start(), m.end(), ln)
        for m in re.finditer("�", ln):
            them("thay_the", so, m.start(), m.end(), ln)
        for m in _LOI_MA_HOA.finditer(ln):
            them("loi_ma_hoa", so, m.start(), m.end(), ln)
        if any(unicodedata.category(c) == "Mn" and ord(c) < 0x370 for c in ln) and unicodedata.normalize("NFC", ln) != ln:
            m = _TO_HOP.search(ln)
            them("unicode_to_hop", so, max(m.start() - 1, 0), m.end(), ln)
        for m in _TU.finditer(ln):
            w = m.group(0)
            if _CHU_GIONG.search(w) and _LA_LATIN.search(w):
                them("chu_la_giong_latin", so, m.start(), m.end(), ln)
        for m in _TOAN_RONG.finditer(ln):
            them("chu_toan_rong", so, m.start(), m.end(), ln)
        for m in _LIGATURE.finditer(ln):
            them("ligature", so, m.start(), m.end(), ln)
        for m in _VUNG_RIENG.finditer(ln):
            if m.group(0) not in ("", "", ""):
                them("ky_tu_vung_rieng", so, m.start(), m.end(), ln)
        for k, ch in enumerate(ln):
            if _la_emoji(ch) and ch != "️":
                them("emoji", so, k, k + 1, ln)
        for rx in _TRICH_DAN:
            for m in rx.finditer(ln):
                them("trich_dan_ai", so, m.start(), m.end(), ln)
        for m in _URL_THEO_DOI.finditer(ln):
            them("url_theo_doi", so, m.start(), m.end(), ln)
        for rx in _CAU_CHATBOT:
            m = rx.search(ln)
            if m:
                them("cau_dan_chatbot", so, m.start(), m.end(), ln)
                break
        for rx in _GIU_CHO:
            for m in rx.finditer(ln):
                them("giu_cho", so, m.start(), m.end(), ln)
        for m in _CHO_KY_DAU.finditer(ln):
            them("cho_ky_dau", so, m.start(), m.end(), ln)
        for m in _GACH_DAI.finditer(ln):
            them("gach_dai", so, m.start(), m.end(), ln)
        n_ngoac_cong += len(_NGOAC_CONG.findall(ln))
        n_ngoac_thang += len(_NGOAC_THANG.findall(ln))
        for m in _SO_MY.finditer(ln):
            them("so_kieu_my", so, m.start(), m.end(), ln)
        # khoảng trắng / dấu câu
        m = re.search(r"\S( {2,})\S", ln) or re.search(r"[ \t]+$", ln)
        if m:
            g = 1 if m.re.pattern.startswith(r"\S") else 0
            them("khoang_trang_thua", so, m.start(g), m.end(g), ln)
        m = _DAU_CAU_LAP.search(ln) or _CACH_TRUOC_DAU.search(ln)
        if m:
            them("dau_cau_sai", so, m.start(), m.end(), ln)
        _quet_markdown(ln, so, che_do, them)
    if n_ngoac_cong and n_ngoac_thang and len(ds) < toi_da:
        ds.append(PhatHien("ngoac_kep_lan", LOAI["ngoac_kep_lan"][0], 0, 0,
                           f"{n_ngoac_thang} dấu thẳng, {n_ngoac_cong} dấu cong", LOAI["ngoac_kep_lan"][1],
                           False, trang))
    return ds


def _quet_markdown(ln, so, che_do, them):
    if che_do == "form":
        m = _MD_TIEUDE_3.match(ln)                          # app chỉ dùng "# " và "## "
        if m:
            them("markdown", so, 0, m.end(1), ln)
    else:
        m = _MD_TIEUDE.match(ln)
        if m:
            them("markdown", so, 0, m.end(1), ln)
        if _MD_GACH_SAO.match(ln):
            m = _MD_GACH_SAO.match(ln)
            them("markdown", so, m.start(), m.end(), ln)
        if _MD_BANG_PHAN_CACH.match(ln) and "|" in ln:
            them("markdown", so, 0, len(ln.strip()), ln)
    if _MD_NGANG.match(ln):
        them("markdown", so, 0, len(ln.strip()), ln)
    for rx in (_MD_DAM, _MD_CODE, _MD_LIEN_KET):
        for m in rx.finditer(ln):
            them("markdown", so, m.start(), m.end(), ln)
    if che_do != "form":
        for m in _MD_NGHIENG.finditer(ln):
            them("markdown", so, m.start(), m.end(), ln)


# ------------------------------------------------------------------ dọn
def _sua_ma_hoa(ln: str) -> str:
    if not _LOI_MA_HOA.search(ln):
        return ln
    for enc in ("cp1252", "latin-1"):
        try:
            f = ln.encode(enc).decode("utf-8")
        except (UnicodeError, ValueError):
            continue
        if len(_LOI_MA_HOA.findall(f)) < len(_LOI_MA_HOA.findall(ln)):
            return f
    return ln


def _sua_dong(ln: str, tc: dict, che_do: str, tk: dict) -> str:
    def dem(ma, n=1):
        if n:
            tk[ma] = tk.get(ma, 0) + n

    if tc["sua_loi_ma_hoa"]:
        moi = _sua_ma_hoa(ln)
        if moi != ln:
            dem("loi_ma_hoa", len(_LOI_MA_HOA.findall(ln)))
            ln = moi
    if tc["go_trich_dan_ai"]:
        for rx in _TRICH_DAN:
            ln, n = rx.subn("", ln)
            dem("trich_dan_ai", n)
    if tc["go_url_theo_doi"]:
        ln, n = _URL_THEO_DOI.subn("", ln)
        dem("url_theo_doi", n)
    if tc["xoa_vo_hinh"]:
        ln, n = _VO_HINH.subn("", ln)
        dem("vo_hinh", n)
    if tc["xoa_dieu_khien"]:
        ln, n = _DIEU_KHIEN.subn("", ln)
        dem("dieu_khien", n)
    if tc["doi_khoang_trang"]:
        ln, n = _KHOANG_TRANG.subn(" ", ln)
        dem("khoang_trang_la", n)
    if tc["go_ligature"]:
        ln, n = _LIGATURE.subn(lambda m: unicodedata.normalize("NFKC", m.group(0)), ln)
        dem("ligature", n)
    if tc["chuan_chu_toan_rong"]:
        ln, n = _TOAN_RONG.subn(lambda m: unicodedata.normalize("NFKC", m.group(0)), ln)
        dem("chu_toan_rong", n)
    if tc["sua_chu_la"]:
        def doi_tu(m):
            w = m.group(0)
            if _CHU_GIONG.search(w) and _LA_LATIN.search(w):
                dem("chu_la_giong_latin", len(_CHU_GIONG.findall(w)))
                return w.translate(_DICH_GIONG)
            return w
        ln = _TU.sub(doi_tu, ln)
    if tc["go_emoji"]:
        da_bo = [0]

        def bo_emoji(m):
            if m.group(0) in _EMOJI_GIU:
                return m.group(0)
            if m.group(0) != "️":
                da_bo[0] += 1
            return ""
        ln = _EMOJI_CHU.sub(bo_emoji, ln)
        dem("emoji", da_bo[0])
    if tc["go_markdown"]:
        ln = _go_markdown(ln, che_do, dem)
    if tc["doi_gach_dai"]:
        ln, n = _GACH_DAI.subn(" - ", ln)
        dem("gach_dai", n)
    if tc["chuan_unicode_nfc"]:
        moi = unicodedata.normalize("NFC", ln)
        if moi != ln:
            dem("unicode_to_hop", len(_TO_HOP.findall(ln)))
            ln = moi
    if tc["sua_dau_cau"]:
        moi = _DAU_CAU_LAP.sub(lambda m: m.group(0)[0], ln)
        moi = _CACH_TRUOC_DAU.sub(r"\1", moi)
        moi = _THIEU_CACH_SAU.sub(r"\1 ", moi)
        if moi != ln:
            dem("dau_cau_sai")
            ln = moi
    if tc["don_khoang_trang"]:
        moi = re.sub(r"(?<=\S) {2,}(?=\S)", " ", ln).rstrip()
        if moi != ln:
            dem("khoang_trang_thua")
            ln = moi
    return ln


def _go_markdown(ln: str, che_do: str, dem) -> str:
    if _MD_NGANG.match(ln):
        dem("markdown")
        return ""
    if che_do == "form":
        m = _MD_TIEUDE_3.match(ln)
        if m:
            dem("markdown")
            ln = "## " + m.group(2)
    else:
        m = _MD_TIEUDE.match(ln)
        if m:
            dem("markdown")
            ln = m.group(2)
        m = _MD_GACH_SAO.match(ln)
        if m:
            dem("markdown")
            ln = m.group(1) + "- " + ln[m.end():]
        if _MD_BANG_PHAN_CACH.match(ln) and "|" in ln:
            dem("markdown")
            return ""
    n_tong = 0
    for _ in range(2):                                           # **__x__** lồng nhau
        ln, n = _MD_DAM.subn(lambda m: m.group(2), ln)
        n_tong += n
    ln, n = _MD_CODE.subn(lambda m: m.group(1), ln)
    n_tong += n
    ln, n = _MD_LIEN_KET.subn(lambda m: m.group(1), ln)
    n_tong += n
    if che_do != "form":
        ln, n = _MD_NGHIENG.subn(lambda m: m.group(1), ln)
        n_tong += n
    dem("markdown", n_tong)
    return ln


def lam_sach(text: str, tuy_chon: dict | None = None, che_do: str = "goc") -> tuple[str, dict[str, int]]:
    """Dọn văn bản theo `tuy_chon` (mặc định xem MAC_DINH). Trả về (văn bản sạch, {mã loại: số chỗ đã sửa})."""
    tc = _tuy_chon(tuy_chon)
    tk: dict[str, int] = {}
    out: list[str] = []
    for ln in text.split("\n"):
        if tc["go_cau_chatbot"] and any(rx.search(ln) for rx in _CAU_CHATBOT):
            tk["cau_dan_chatbot"] = tk.get("cau_dan_chatbot", 0) + 1
            continue
        moi = _sua_dong(ln.rstrip("\r"), tc, che_do, tk)
        if moi == "" and ln.strip():                             # dòng chỉ chứa đường kẻ markdown: bỏ hẳn
            continue
        out.append(moi)
    return "\n".join(out), tk


def lam_sach_dong(ln: str, tuy_chon: dict | None = None, che_do: str = "goc") -> tuple[str, dict[str, int]]:
    """Dọn một dòng; dòng bị bỏ (đường kẻ markdown) trả về chuỗi rỗng."""
    tc = _tuy_chon(tuy_chon)
    tk: dict[str, int] = {}
    if tc["go_cau_chatbot"] and any(rx.search(ln) for rx in _CAU_CHATBOT):
        return "", {"cau_dan_chatbot": 1}
    return _sua_dong(ln, tc, che_do, tk), tk


# ------------------------------------------------------------------ tóm tắt
def dem_theo_loai(ds: list[PhatHien]) -> dict[str, int]:
    out: dict[str, int] = {}
    for p in ds:
        out[p.ma] = out.get(p.ma, 0) + 1
    return out


def tom_tat(ds: list[PhatHien], da_don: dict[str, int] | None = None, toi_da_dong: int = 8) -> list[str]:
    """Các câu cảnh báo ngắn, mỗi loại một dòng (dùng cho danh sách cảnh báo của ứng dụng)."""
    out = []
    cac = {}
    for p in ds:
        cac.setdefault(p.ma, []).append(p)
    thu_tu = {"cao": 0, "trung_binh": 1, "thap": 2}
    for ma, lst in sorted(cac.items(), key=lambda kv: (thu_tu[LOAI[kv[0]][0]], kv[0])):
        dong = sorted({x.dong for x in lst if x.dong})
        vitri = ""
        if dong:
            vitri = " (dòng " + ", ".join(map(str, dong[:5])) + ("…" if len(dong) > 5 else "") + ")"
        trang = sorted({x.trang for x in lst if x.trang})
        if trang:
            vitri = " (trang " + ", ".join(map(str, trang[:5])) + ("…" if len(trang) > 5 else "") + ")"
        da = ""
        if da_don and da_don.get(ma):
            da = " Đã tự dọn."
        out.append(f"Dò ký tự lạ: {len(lst)} chỗ “{LOAI[ma][1].split(' (')[0]}”{vitri}.{da}")
    if len(out) > toi_da_dong:
        out = out[:toi_da_dong] + [f"Dò ký tự lạ: còn {len(out) - toi_da_dong} loại khác, xem mục “Dò ký tự lạ”."]
    return out


def quet_nhieu_dong(dong: list[str], che_do: str = "goc") -> list[PhatHien]:
    return quet("\n".join(dong), che_do=che_do)
