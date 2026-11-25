# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Đọc file PowerPoint thô thành mô hình slide đơn giản.

Giữ lại: chữ (theo cấp), hình (thu nhỏ nếu quá lớn, sắp theo vị trí trên slide), bảng
(kể cả ô gộp: "<" gộp với ô bên trái, "^" gộp với ô phía trên), ghi chú diễn giả.
Biểu đồ được đổi thành bảng số liệu. SmartArt, video, OLE chỉ được báo để chèn lại.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from ..hinh import EXT, xu_ly


@dataclass
class SlideIn:
    title: str = ""
    items: list = field(default_factory=list)      # (cấp, văn bản, có gạch đầu dòng)
    images: list = field(default_factory=list)     # (bytes, đuôi file)
    tables: list = field(default_factory=list)     # list[list[list[str]]]
    notes: str = ""
    unsupported: list = field(default_factory=list)
    canh_bao: list = field(default_factory=list)   # cảnh báo riêng của slide (hình hỏng, biểu đồ -> bảng...)
    chu_anh: list = field(default_factory=list)    # chữ nằm đè lên hình -> chú thích, song song với images
    so_do: bool = False                            # slide dạng sơ đồ/infographic (nhiều mảnh chữ rời + hình vẽ)
    nguon: object = None                           # slide gốc (để chép nguyên hình vẽ khi slide là sơ đồ)
    bo_el: list = field(default_factory=list)      # phần tử không chép (tiêu đề đã lấy riêng)
    kich_goc: tuple = (0, 0)                       # khổ slide gốc (EMU)
    n_nho: int = 0                                 # đếm để nhận diện sơ đồ
    n_hinh: int = 0
    n_wa: int = 0
    phuc_tap: bool = False                         # có biểu đồ/SmartArt/video: không chép nguyên


_TITLE = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_BO_QUA = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}
_RE_GACH = re.compile(r"^[\-–—•●▪■◦○+*➢➤✓✔]\s*")


def _lam_sach(s: str) -> str:
    s = s.replace("\x0b", " ").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip()


TIEU_DE_TOI_DA = 100      # tiêu đề dài hơn mức này là đoạn văn nằm nhầm trong ô tiêu đề


def _doan_chu(p) -> str:
    """Chữ của một đoạn, GIỮ chỗ xuống dòng mềm (<a:br>, Shift+Enter) thành dấu cách.
    Trước đây ghép các run nên chữ hai bên chỗ xuống dòng dính nhau (\"SprinklerĐầu\")."""
    try:
        t = p.text                      # python-pptx đổi <a:br> thành \x0b
    except Exception:
        t = "".join(r.text for r in p.runs)
    return _lam_sach(t.replace("\x0b", " ").replace("\n", " "))


def _doan_theo_dong(p) -> list[str]:
    """Như _doan_chu nhưng tách tại mỗi chỗ xuống dòng mềm (dùng cho ô tiêu đề)."""
    try:
        t = p.text
    except Exception:
        t = "".join(r.text for r in p.runs)
    return [x for x in (_lam_sach(k) for k in re.split(r"[\x0b\n]", t)) if x]


def tach_tieu_de(text: str) -> tuple[str, str]:
    """Tách một đoạn chữ dài thành (tiêu đề ngắn, phần còn lại cho nội dung).
    Đoạn ngắn thì giữ nguyên làm tiêu đề. Không bao giờ làm mất chữ."""
    text = _lam_sach(text)
    if len(text) <= TIEU_DE_TOI_DA:
        return text, ""
    m = re.match(r"^(.{12,%d}?[.:;!?])\s+(\S.*)$" % TIEU_DE_TOI_DA, text)
    if m:
        return m.group(1).rstrip(".:;").strip(), m.group(2).strip()
    cat = text.rfind(" ", 0, TIEU_DE_TOI_DA - 10)
    cat = cat if cat > 20 else TIEU_DE_TOI_DA - 10
    return text[:cat].rstrip(" ,;:-") + "…", text      # giữ nguyên toàn văn ở nội dung


def _bang(table) -> list[list[str]]:
    """Bảng PowerPoint -> lưới chữ; ô gộp đánh dấu '<' (gộp trái) và '^' (gộp trên)."""
    nr, nc = len(table.rows), len(table.columns)
    luoi = [["" for _ in range(nc)] for _ in range(nr)]
    for i in range(nr):
        for j in range(nc):
            c = table.cell(i, j)
            if c.is_merge_origin:
                luoi[i][j] = _lam_sach(c.text)
                for di in range(c.span_height):
                    for dj in range(c.span_width):
                        if di or dj:
                            if i + di < nr and j + dj < nc:
                                luoi[i + di][j + dj] = "^" if (dj == 0 and di > 0) else "<"
            elif not c.is_spanned:
                luoi[i][j] = _lam_sach(c.text)
    return luoi


def _bang_tu_bieu_do(chart) -> list[list[str]] | None:
    try:
        plot = chart.plots[0]
        cats = [str(c) for c in plot.categories]
        series = list(plot.series)
        if not cats or not series:
            return None
        rows = [[""] + [str(s.name or f"Dãy {k + 1}") for k, s in enumerate(series)]]
        for i, c in enumerate(cats):
            row = [c]
            for s in series:
                vals = list(s.values)
                v = vals[i] if i < len(vals) else ""
                row.append("" if v is None else (f"{v:g}" if isinstance(v, (int, float)) else str(v)))
            rows.append(row)
        return rows
    except Exception:
        return None


def _la_smartart(sh) -> bool:
    try:
        return any("diagram" in (u or "") for u in
                   sh._element.xpath(".//a:graphicData/@uri"))
    except Exception:
        return False



def _la_media(sh) -> bool:
    try:
        x = sh._element
        return bool(x.xpath(".//a:videoFile | .//a:audioFile | .//a:wavAudioFile | .//*[local-name()='media']"))
    except Exception:
        return False


def _chu_smartart(sh) -> list[str]:
    """Lấy các dòng chữ trong phần dữ liệu của SmartArt (dgm:pt loại node)."""
    try:
        rid = sh._element.xpath(".//*[local-name()='relIds']/@*[local-name()='dm']")[0]
        part = sh.part.related_part(rid)
        from lxml import etree
        root = etree.fromstring(part.blob)
        ns = {"dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
              "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        out = []
        for pt in root.findall(".//dgm:pt", ns):
            if pt.get("type", "node") != "node":
                continue
            t = " ".join("".join(p.itertext()).strip() for p in pt.findall(".//a:p", ns)).strip()
            if t:
                out.append(t)
        return out
    except Exception:
        return []


def _cat_anh(sh, data: bytes, mime: str):
    """Áp dụng vùng cắt (crop) của ảnh gốc để slide mới không hiện phần đã bị cắt."""
    try:
        l, t, r, b = sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom
        if not any((l, t, r, b)) or l + r >= 0.95 or t + b >= 0.95:
            return data, mime
        import io
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        w, h = im.size
        im = im.crop((int(w * l), int(h * t), int(w * (1 - r)), int(h * (1 - b))))
        buf = io.BytesIO()
        fmt = "PNG" if mime == "image/png" else "JPEG"
        if fmt == "JPEG":
            im = im.convert("RGB")
        im.save(buf, fmt)
        return buf.getvalue(), mime
    except Exception:
        return data, mime

_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _mo_alternate(slide):
    """python-pptx bỏ qua mọi hình nằm trong <mc:AlternateContent> (WordArt, hiệu ứng, chữ đặc biệt...),
    nên chữ trong đó bị mất. Mở các khối này ra (lấy nhánh Choice, không có thì Fallback) trước khi đọc."""
    tree = slide.shapes._spTree
    for ac in list(tree.iter("{%s}AlternateContent" % _MC)):
        cha = ac.getparent()
        if cha is None:
            continue
        chon = ac.find("{%s}Choice" % _MC)
        du = ac.find("{%s}Fallback" % _MC)
        nguon = chon if (chon is not None and len(chon)) else du
        vt = cha.index(ac)
        con = list(nguon) if nguon is not None else []
        cha.remove(ac)
        for k, e in enumerate(con):
            cha.insert(vt + k, e)


def _hop(sh, cha):
    try:
        t, l, w, h = sh.top or 0, sh.left or 0, sh.width or 0, sh.height or 0
    except Exception:
        t = l = w = h = 0
    if cha:
        t, l = cha
    return (t, l), (w, h)


def _duyet(shapes, sl: SlideIn, khoi: list, tieu_de: list, anh: list, cha=None):
    for sh in shapes:
        try:
            st = sh.shape_type
        except NotImplementedError:
            st = None
        vt, kt = _hop(sh, cha)
        if st == MSO_SHAPE_TYPE.GROUP:
            _duyet(sh.shapes, sl, khoi, tieu_de, anh, vt)
            continue
        try:
            if sh._element.xpath(".//a:prstTxWarp[@prst and @prst!='textNoShape']"):
                sl.n_wa += 1
        except Exception:
            pass
        if getattr(sh, "has_chart", False) and sh.has_chart:
            sl.phuc_tap = True
            rows = _bang_tu_bieu_do(sh.chart)
            if rows:
                sl.tables.append(rows)
                sl.canh_bao.append("có biểu đồ, đã đổi thành bảng số liệu; hãy vẽ lại biểu đồ nếu cần")
            else:
                sl.unsupported.append("biểu đồ")
            continue
        if getattr(sh, "has_table", False) and sh.has_table:
            sl.tables.append(_bang(sh.table))
            continue
        if _la_smartart(sh):
            sl.phuc_tap = True
            chu = _chu_smartart(sh)
            if chu:
                sl.items.extend((0, t, True) for t in chu)
                sl.canh_bao.append("có SmartArt, đã lấy phần chữ thành các ý; hãy vẽ lại sơ đồ nếu cần")
            else:
                sl.unsupported.append("SmartArt")
            continue
        if _la_media(sh):                       # hình poster của video/âm thanh: thường là khung đen
            sl.phuc_tap = True
            sl.unsupported.append("đối tượng nhúng (video/âm thanh/OLE)")
            continue
        if st == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and hasattr(sh, "image")):
            try:
                blob = sh.image.blob
            except Exception:
                sl.canh_bao.append("có hình liên kết tới file ngoài hoặc bị hỏng, chưa chuyển được; "
                                   "hãy chèn lại hình")
                continue
            r = xu_ly(blob)
            if r is None:
                sl.canh_bao.append("có hình định dạng chưa đọc được (thường là EMF/WMF); hãy chèn lại "
                                   "hình dạng PNG/JPG")
                continue
            data, mime, _, _ = r
            data, mime = _cat_anh(sh, data, mime)
            anh.append((vt, (data, EXT[mime]), kt))
            continue
        if st in (MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
                  MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
            sl.phuc_tap = True
            sl.unsupported.append("đối tượng nhúng (video/âm thanh/OLE)")
            continue
        co_chu = sh.has_text_frame and bool(sh.text_frame.text.strip())
        if not co_chu:
            if not sh.is_placeholder and (st in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.FREEFORM,
                                                  MSO_SHAPE_TYPE.AUTO_SHAPE) or sh._element.tag.endswith("cxnSp")):
                sl.n_hinh += 1
            continue
        ph = sh.placeholder_format.type if sh.is_placeholder else None
        if ph in _BO_QUA:
            continue
        if (sh.name or "").startswith(("vbn-chan-trang", "vbn-so-trang")):
            continue                                        # chân trang do chính ứng dụng dựng
        paras = []
        if ph in _TITLE:
            # ô tiêu đề: mỗi dòng (kể cả xuống dòng mềm) là một mảnh riêng để còn tách tiêu đề khỏi nội dung
            for p in sh.text_frame.paragraphs:
                paras.extend((0, t, False) for t in _doan_theo_dong(p))
        else:
            for p in sh.text_frame.paragraphs:
                t = _doan_chu(p)
                if not t:
                    continue
                gach = bool(_RE_GACH.match(t))
                t = _RE_GACH.sub("", t)
                la_than = sh.is_placeholder and ph not in _TITLE
                paras.append((min(p.level, 1), t, gach or la_than or p.level > 0))
        if not paras:
            continue
        if ph in _TITLE:
            tieu_de.append(([t for _, t, _ in paras], sh._element))
        else:
            if not sh.is_placeholder and sum(len(t.split()) for _, t, _ in paras) <= 4:
                sl.n_nho += 1
            khoi.append((vt, paras, kt, sh._element))


def _hang(khoi: list) -> list:
    """Gom các mảnh chữ rời (sơ đồ) thành từng hàng theo vị trí: cùng hàng thì ghép cạnh nhau."""
    kk = sorted(khoi, key=lambda x: (x[0][0], x[0][1]))
    hang, top0 = [], None
    for vt, paras, _kt, _el in kk:
        txt = " ".join(t for _, t, _ in paras)
        if top0 is None or vt[0] - top0 > 274320:         # 0,3 inch
            hang.append([(vt[1], txt)])
            top0 = vt[0]
        else:
            hang[-1].append((vt[1], txt))
    return [(0, "   ".join(t for _, t in sorted(h)), True) for h in hang]


def doc_pptx(data: bytes) -> tuple[list[SlideIn], list[str]]:
    prs = Presentation(io.BytesIO(data))
    W, H = int(prs.slide_width), int(prs.slide_height)
    out: list[SlideIn] = []
    warns: list[str] = []
    for idx, slide in enumerate(prs.slides, 1):
        _mo_alternate(slide)
        sl = SlideIn()
        khoi, tieu_de, anh = [], [], []
        _duyet(slide.shapes, sl, khoi, tieu_de, anh)
        sl.so_do = (not sl.phuc_tap and not sl.tables and
                    (sl.n_nho >= 4 or (sl.n_hinh >= 2 and sl.n_nho >= 2) or sl.n_wa >= 2))
        chu_anh = [""] * len(anh)
        if not sl.so_do and anh:                           # chữ đè lên hình -> chú thích của hình đó
            con = []
            for kh in khoi:
                (t, l), (w, h), tx = kh[0], kh[2], " ".join(x[1] for x in kh[1])
                cx, cy = l + w // 2, t + h // 2
                j = next((k for k, a in enumerate(anh)
                          if a[2][0] and a[2][1] and a[0][1] <= cx <= a[0][1] + a[2][0]
                          and a[0][0] <= cy <= a[0][0] + a[2][1]), None)
                if j is not None and len(tx.split()) <= 6 and len(tx) <= 60:
                    chu_anh[j] = (chu_anh[j] + " " + tx).strip()
                else:
                    con.append(kh)
            khoi = con
        khoi.sort(key=lambda x: (x[0][0], x[0][1]))
        xep = sorted(zip(anh, chu_anh), key=lambda z: (z[0][0][1] // 457200 * 0 + z[0][0][0] // 914400, z[0][0][1]))
        sl.images = [a[1] for a, _ in xep]
        sl.chu_anh = [c for _, c in xep]
        phan_them: list = []                              # chữ nằm nhầm trong ô tiêu đề, đưa xuống nội dung
        bo_el = None
        if tieu_de:
            manh, bo_el = tieu_de[0]
            sl.title, du = tach_tieu_de(manh[0])
            if du:
                phan_them.append((0, du, False))
            phan_them.extend((0, t, False) for t in manh[1:])
            if phan_them:
                sl.canh_bao.append("ô tiêu đề chứa cả đoạn văn dài hoặc nhiều dòng; đã tách dòng đầu làm "
                                   "tiêu đề, phần còn lại đưa xuống nội dung")
            if phan_them:
                bo_el = None
        elif khoi:                                        # không có placeholder tiêu đề
            first = khoi[0][1]
            tt, du = tach_tieu_de(first[0][1])
            sl.title = tt
            if len(first[0][1]) <= TIEU_DE_TOI_DA:
                khoi[0] = (khoi[0][0], first[1:], khoi[0][2], khoi[0][3])
                if not first[1:]:
                    bo_el = khoi[0][3]
            elif du:
                khoi[0] = (khoi[0][0], [(first[0][0], du, first[0][2])] + first[1:], khoi[0][2], khoi[0][3])
        sl.items.extend(phan_them)
        if sl.so_do:
            sl.items.extend(_hang([k for k in khoi if k[1]]))
            sl.nguon, sl.kich_goc = slide, (W, H)
            sl.bo_el = [bo_el] if bo_el is not None else []
            sl.canh_bao.append("slide dạng sơ đồ/hình vẽ (nhiều mảnh chữ rời); giữ nguyên hình vẽ gốc, chỉ "
                               "chuẩn hóa tiêu đề và chân trang")
        else:
            for _, paras, _kt, _el in khoi:
                sl.items.extend(paras)
        if slide.has_notes_slide:
            sl.notes = slide.notes_slide.notes_text_frame.text.strip()
        if sl.unsupported:
            warns.append(f"Slide {idx}: có {', '.join(sorted(set(sl.unsupported)))} chưa chuyển "
                         "được, hãy chèn lại thủ công.")
        for w in dict.fromkeys(sl.canh_bao):
            warns.append(f"Slide {idx}: {w}.")
        if not sl.title and (sl.items or sl.images or sl.tables):
            warns.append(f"Slide {idx}: không xác định được tiêu đề, đã để “Slide {idx}”.")
            sl.title = f"Slide {idx}"
        if sl.title or sl.items or sl.images or sl.tables:
            out.append(sl)
    return out, warns
