# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Dựng lại bộ slide theo một chuẩn trình bày (profile): tiêu chuẩn chuyên nghiệp hoặc học thuật.
(Tên file giữ nguyên để tương thích; kiểu mặc định hiện là \"slide_chuan\".)"""
from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .. import ban_quyen
from ..importers.pptx_reader import SlideIn
from ..profile import load_profile
from .common import tinh_do_rong_cot

RE_TLTK = re.compile(r"tài liệu tham khảo|references|bibliography", re.I)
RE_SO = re.compile(r"(\d[\d.,]*\s?(?:%|tỷ|triệu|nghìn|ngàn|lần|đồng|VND|USD|\$)?)")
RE_CAM_ON = re.compile(r"cảm ơn|thank|hỏi\s*đáp|q\s*&\s*a|câu hỏi|thảo luận", re.I)

KIEU_SLIDE = {"chuan": "slide_chuan", "hoc_thuat": "slide_hoc_thuat"}     # tên gọi ngoài giao diện -> profile
KIEU_MAC_DINH = "chuan"


def ten_profile(kieu: str | None) -> str:
    """Đổi tên kiểu người dùng chọn thành tên profile (bí danh cũ hoặc bất kỳ profile loại slide);
    không hợp lệ thì dùng kiểu mặc định."""
    k = (kieu or "").strip().lower()
    if k in KIEU_SLIDE:
        return KIEU_SLIDE[k]
    from ..profile import list_profiles
    return k if k in list_profiles("slide") else KIEU_SLIDE[KIEU_MAC_DINH]


@dataclass
class ThongTin:
    ten_bai: str = ""
    hoc_phan: str = ""
    giang_vien: str = ""
    don_vi: str = ""
    nam: str = ""


def _rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h)


class SlideBuilder:
    def __init__(self, prof: dict, info: ThongTin):
        self.prof, self.info = prof, info
        self.font = prof["font"]
        self.c = prof["colors"]
        self._ap_vung()
        self.prs = Presentation()
        self.prs.slide_width = Inches(prof["slide"]["width_in"])
        self.prs.slide_height = Inches(prof["slide"]["height_in"])
        self.warns: list[str] = []
        self.prs.core_properties.comments = ban_quyen.GHI_CHU

    # ---------- tiện ích ----------
    def _ap_vung(self):
        """Tính lại vị trí tiêu đề/nội dung/chân trang từ các tỷ lệ % (layout.use_percent)."""
        pr = self.prof
        ly = pr.get("layout") or {}
        self.vung_chan = None
        if not ly.get("use_percent"):
            return
        W, H = pr["slide"]["width_in"], pr["slide"]["height_in"]
        x = round(W * ly["side_margin_pct"] / 100, 3)
        y0 = round(H * ly["vertical_margin_pct"] / 100, 3)
        tz = H * ly["title_zone_pct"] / 100
        cz = H * ly["center_zone_pct"] / 100
        t, b = pr["title"], pr["body"]
        t["left_in"], t["width_in"], t["top_in"] = x, round(W - 2 * x, 3), y0
        t["height_in"] = round(max(tz - y0 - 0.1, 0.6), 3)
        b["left_in"], b["width_in"] = x, round(W - 2 * x, 3)
        b["top_in"] = round(tz + 0.1, 3)
        b["height_in"] = round(cz - 0.2, 3)
        self.vung_chan = round(H * (1 - ly["footer_zone_pct"] / 100), 3)
        if "width_with_image_in" in b:
            b["width_with_image_in"] = round((W - 2 * x) * 0.56, 3)
        im = pr["image"]
        im["right_x_in"] = round(x + b["width_with_image_in"] + 0.3, 3)
        im["right_w_in"] = round(W - x - im["right_x_in"], 3)

    def _style_run(self, run, size, bold=False, color=None, italic=False):
        f = run.font
        f.name, f.size, f.bold, f.italic = self.font, Pt(size), bold, italic
        f.color.rgb = _rgb(color or self.c["text"])

    def _bullet(self, p, level: int, kind: str):
        """kind: 'dot' | 'num' | 'none'."""
        pPr = p._p.get_or_add_pPr()
        marl = {0: 0.38, 1: 0.85}[level] if kind != "none" else 0
        ind = -0.32 if level == 0 else -0.27
        pPr.set("marL", str(int(Inches(marl))))
        pPr.set("indent", str(int(Inches(ind))) if kind != "none" else "0")
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        if kind == "none":
            etree.SubElement(pPr, qn("a:buNone"))
        elif kind == "num":
            self._mau_dau_dong(pPr)
            el = etree.SubElement(pPr, qn("a:buAutoNum"))
            el.set("type", "arabicPeriod")
        else:
            self._mau_dau_dong(pPr)
            bf = etree.SubElement(pPr, qn("a:buFont"))
            bf.set("typeface", "Arial")
            bc = etree.SubElement(pPr, qn("a:buChar"))
            bc.set("char", "•" if level == 0 else "–")

    def _mau_dau_dong(self, pPr):
        """Dấu đầu dòng/số thứ tự tô màu điểm nhấn nếu profile yêu cầu (phải đứng trước buFont/buChar)."""
        if self.prof["body"].get("bullet_color") == "accent":
            clr = etree.SubElement(pPr, qn("a:buClr"))
            etree.SubElement(clr, qn("a:srgbClr")).set("val", self.c["accent"])

    def _duong_ke(self, slide, x1, y1, x2, y2, color, pt):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        ln.line.color.rgb = _rgb(color)
        ln.line.width = Pt(pt)
        return ln

    def _khoi(self, slide, x, y, w, h, color):
        r = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = _rgb(color)
        r.line.fill.background()
        return r

    def _footer(self, slide, so_slide=True):
        pr, ft = self.prof, self.prof["footer"]
        W, H = pr["slide"]["width_in"], pr["slide"]["height_in"]
        text = self.info.don_vi or self.info.ten_bai
        xl = pr["title"]["left_in"] if self.vung_chan else 0.6
        yl = (self.vung_chan + 0.1) if self.vung_chan else H - 0.62
        if ft.get("line"):
            self._duong_ke(slide, xl, yl, W - xl, yl, self.c.get("border", self.c["muted"]), 0.75)
        if text:
            tb = slide.shapes.add_textbox(Inches(xl), Inches(yl + 0.07), Inches(9), Inches(0.35))
            tb.text_frame.word_wrap = True
            tb.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = text
            self._style_run(r, ft["size_pt"], color=self.c["muted"])
        if so_slide:
            tb = slide.shapes.add_textbox(Inches(W - xl - 1.0), Inches(yl + 0.07), Inches(1.0), Inches(0.35))
            p = tb.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            fld = etree.SubElement(p._p, qn("a:fld"))
            fld.set("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}")
            fld.set("type", "slidenum")
            rpr = etree.SubElement(fld, qn("a:rPr"))
            rpr.set("lang", "vi-VN")
            rpr.set("sz", str(ft["size_pt"] * 100))
            sf = etree.SubElement(rpr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", self.c["muted"])
            etree.SubElement(rpr, qn("a:latin")).set("typeface", self.font)
            etree.SubElement(fld, qn("a:t")).text = "‹#›"

    # ---------- các loại slide ----------
    def _dong_thong_tin(self) -> list[str]:
        """Các dòng thông tin dưới tên báo cáo/bài giảng, kèm nhãn theo profile."""
        nhan = self.prof["title_slide"].get("nhan", {})
        mac_dinh = {"hoc_phan": "Học phần", "giang_vien": "Giảng viên"}

        def gan(khoa, gia_tri):
            if not gia_tri:
                return ""
            lb = nhan.get(khoa, mac_dinh.get(khoa, ""))
            return f"{lb}: {gia_tri}" if lb else gia_tri
        return [x for x in (gan("hoc_phan", self.info.hoc_phan), gan("giang_vien", self.info.giang_vien),
                            self.info.don_vi, self.info.nam) if x]

    def _nen_toi(self, s):
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = _rgb(self.c["primary"])

    def slide_tieu_de(self, ten_bai: str):
        pr, ts = self.prof, self.prof["title_slide"]
        trai = ts.get("layout", "giua") == "trai"
        s = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        self._nen_toi(s)
        W, H = pr["slide"]["width_in"], pr["slide"]["height_in"]
        if trai:
            if self.c.get("panel"):                                   # khối trang trí bên phải
                self._khoi(s, W - 3.4, 0, 3.4, H, self.c["panel"])
            self._khoi(s, 0.9, 2.05, 0.09, 1.75, self.c["accent"])
            x, w = 1.25, W - 1.25 - 3.9
        else:
            x, w = 1.0, W - 2.0
        t = s.shapes.title
        t.left, t.top, t.width, t.height = Inches(x), Inches(2.0), Inches(w), Inches(1.9)
        tf = t.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT if trai else PP_ALIGN.CENTER
        r = tf.paragraphs[0].add_run()
        r.text = ten_bai.upper() if ts.get("uppercase", True) else ten_bai
        size = ts["title_size_pt"]
        cpl = pr["title"]["chars_per_line_at_32"] * 32 / size * (w / (W - 2.0))
        while size > 28 and len(r.text) > cpl * 3:                    # tối đa ~3 dòng trong khung
            size -= 2
            cpl = pr["title"]["chars_per_line_at_32"] * 32 / size * (w / (W - 2.0))
        self._style_run(r, size, True, self.c["on_primary"])
        sub = s.placeholders[1]
        sub.left, sub.top, sub.width, sub.height = Inches(x), Inches(4.15), Inches(w), Inches(2.3)
        lines = self._dong_thong_tin()
        stf = sub.text_frame
        stf.word_wrap = True
        for i, ln in enumerate(lines):
            p = stf.paragraphs[0] if i == 0 else stf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT if trai else PP_ALIGN.CENTER
            p.space_after = Pt(4)
            self._bullet(p, 0, "none")
            rr = p.add_run()
            rr.text = ln
            self._style_run(rr, ts["sub_size_pt"], False, self.c["on_primary"])
        if not lines:
            sub._element.getparent().remove(sub._element)

    def slide_ket(self, text: str = "", dong_phu: list[str] | None = None):
        """Slide kết (cảm ơn / hỏi đáp), cùng bộ nhận diện với slide mở đầu."""
        cl, ts = self.prof["closing"], self.prof["title_slide"]
        trai = ts.get("layout", "giua") == "trai"
        W, H = self.prof["slide"]["width_in"], self.prof["slide"]["height_in"]
        s = self.prs.slides.add_slide(self.prs.slide_layouts[5])
        self._nen_toi(s)
        if trai and self.c.get("panel"):
            self._khoi(s, W - 3.4, 0, 3.4, H, self.c["panel"])
        if trai:
            self._khoi(s, 0.9, 2.55, 0.09, 1.3, self.c["accent"])
        x, w = (1.25, W - 1.25 - 3.9) if trai else (1.0, W - 2.0)
        t = s.shapes.title
        t.left, t.top, t.width, t.height = Inches(x), Inches(2.4), Inches(w), Inches(1.5)
        tf = t.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT if trai else PP_ALIGN.CENTER
        r = tf.paragraphs[0].add_run()
        r.text = text or cl["text"]
        self._style_run(r, cl["size_pt"], True, self.c["on_primary"])
        if dong_phu:
            tb = s.shapes.add_textbox(Inches(x), Inches(4.0), Inches(w), Inches(2.0))
            tb.text_frame.word_wrap = True
            for i, ln in enumerate(dong_phu):
                p = tb.text_frame.paragraphs[0] if i == 0 else tb.text_frame.add_paragraph()
                p.alignment = PP_ALIGN.LEFT if trai else PP_ALIGN.CENTER
                rr = p.add_run()
                rr.text = ln
                self._style_run(rr, ts["sub_size_pt"], False, self.c["on_primary"])

    def _tao_slide(self, tieu_de: str):
        pr, tp = self.prof, self.prof["title"]
        s = self.prs.slides.add_slide(self.prs.slide_layouts[5])   # Title Only
        W = pr["slide"]["width_in"]
        hd = pr.get("header", {})
        if hd.get("top_bar", True):
            bar = s.shapes.add_shape(1, 0, 0, self.prs.slide_width, Inches(0.12))
            bar.fill.solid()
            bar.fill.fore_color.rgb = _rgb(self.c["primary"])
            bar.line.fill.background()
        if hd.get("marker"):                                          # khối nhấn màu bên trái tiêu đề
            self._khoi(s, tp["left_in"] - 0.3, tp["top_in"] + 0.17, 0.09, tp["height_in"] - 0.34, self.c["accent"])
        t = s.shapes.title
        t.left, t.top = Inches(tp["left_in"]), Inches(tp["top_in"])
        t.width, t.height = Inches(tp["width_in"]), Inches(tp["height_in"])
        tf = t.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        size = tp["size_pt"]
        cpl = tp["chars_per_line_at_32"]
        while size > tp["min_size_pt"] and len(tieu_de) > cpl * 32 / size * 1.0:
            size -= 2
        toi_da = int(cpl * 32 / size * tp.get("max_lines", 2))     # tiêu đề không được tràn ra ngoài khung
        if len(tieu_de) > toi_da:
            cat = tieu_de.rfind(" ", 0, toi_da - 1)
            tieu_de_goc = tieu_de
            tieu_de = tieu_de[:cat if cat > 10 else toi_da - 1].rstrip(" ,;:-") + "…"
            self.warns.append(f"Tiêu đề quá dài, đã rút gọn để không tràn slide: “{tieu_de_goc[:60]}…”.")
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT
        r = tf.paragraphs[0].add_run()
        r.text = tieu_de
        self._style_run(r, size, True, self.c["primary"])
        kieu_ke = hd.get("underline", "dai")
        if kieu_ke != "khong":
            y_ke = tp["top_in"] + tp["height_in"] + 0.05
            mau_ke = self.c["accent"] if kieu_ke == "dai" else self.c.get("border", self.c["accent"])
            self._duong_ke(s, tp["left_in"], y_ke, W - tp["left_in"], y_ke, mau_ke, hd.get("underline_pt", 2))
        self._footer(s)
        return s

    def _body_box(self, s, width_in):
        b = self.prof["body"]
        tb = s.shapes.add_textbox(Inches(b["left_in"]), Inches(b["top_in"]),
                                  Inches(width_in), Inches(b["height_in"]))
        tf = tb.text_frame
        tf.word_wrap = True
        return tf

    def _uoc_luong_dong(self, text, level, width_in):
        b = self.prof["body"]
        size = b["size_pt"] if level == 0 else b["sub_size_pt"]
        cpl = max(int((width_in - 0.6) * 72 / (size * 0.47)), 10)
        return max(1, math.ceil(len(text) / cpl))

    def _dien_gach(self, tf, items, width_in):
        b = self.prof["body"]
        first = True
        for level, text, gach in items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_after = Pt(8)
            p.line_spacing = 1.05
            self._bullet(p, level, "dot" if gach else "none")
            size = b["size_pt"] if level == 0 else b["sub_size_pt"]
            if b.get("highlight_numbers") and self.c.get("highlight"):
                for kt, doan in enumerate(RE_SO.split(text)):
                    if not doan:
                        continue
                    r = p.add_run()
                    r.text = doan
                    if kt % 2 == 1:
                        self._style_run(r, size, True, self.c["highlight"])
                    else:
                        self._style_run(r, size)
            else:
                r = p.add_run()
                r.text = text
                self._style_run(r, size)

    def _rut_y(self, sl):
        """long_item = notes: giữ ý chính trên slide, câu đầy đủ chuyển vào ghi chú."""
        ck = self.prof.get("standard", {}).get("chunking", {})
        if ck.get("long_item") != "notes":
            return sl.items
        mx = ck.get("keyword_words", 12)
        moi, bo = [], []
        for lv, tx, g in sl.items:
            tu = tx.split()
            if len(tu) > mx:
                bo.append(tx)
                tx = " ".join(tu[:mx]).rstrip(" ,;:-") + "…"
            moi.append((lv, tx, g))
        if bo:
            sl.notes = (sl.notes + "\n\n" if sl.notes else "") + "Nội dung đầy đủ:\n" + "\n".join(bo)
            self.warns.append(f"Slide “{sl.title}”: {len(bo)} ý dài đã rút thành ý chính, câu đầy đủ nằm trong ghi chú.")
        return moi

    def _la_the(self, sl) -> bool:
        cl = self.prof.get("content_layout") or {}
        if not cl.get("cards") or sl.images or sl.tables:
            return False
        it = sl.items
        if not (cl.get("min_items", 3) <= len(it) <= 5) or any(x[0] != 0 for x in it):
            return False
        return all(len(x[1].split()) <= cl.get("card_max_words", 14) for x in it)

    def _ve_the(self, s, items):
        b = self.prof["body"]
        n = len(items)
        gap = 0.25
        w = (b["width_in"] - gap * (n - 1)) / n
        h = min(b["height_in"] - 0.6, 3.4)
        y = b["top_in"] + 0.3
        for i, (_, tx, _g) in enumerate(items):
            x = b["left_in"] + i * (w + gap)
            self._khoi(s, x, y, w, h, self.c.get("card", "F8F9FA"))
            self._khoi(s, x, y, w, 0.08, self.c["accent"])
            tb = s.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.25), Inches(w - 0.3), Inches(h - 0.4))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = f"{i + 1:02d}"
            self._style_run(r, 28, True, self.c["highlight"] if self.c.get("highlight") else self.c["accent"])
            p2 = tf.add_paragraph()
            p2.space_before = Pt(8)
            r2 = p2.add_run()
            r2.text = tx
            self._style_run(r2, max(b["sub_size_pt"], 18))

    def _chia_doan_dai(self, items, width_in):
        """Đoạn dài hơn số dòng tối đa của một slide thì tách theo câu (không làm mất chữ),
        phần sau sang slide \"(tiếp)\"; trước đây một đoạn quá dài làm chữ tràn khỏi slide."""
        b = self.prof["body"]
        toi_da = b["max_lines"]
        out = []
        for level, text, gach in items:
            if self._uoc_luong_dong(text, level, width_in) <= toi_da:
                out.append((level, text, gach))
                continue
            cau = [x for x in re.split(r"(?<=[.!?;:…])\s+", text) if x]
            manh, cur = [], ""
            for c in cau:
                cand = f"{cur} {c}".strip()
                if self._uoc_luong_dong(cand, level, width_in) <= toi_da:
                    cur = cand
                    continue
                if cur:
                    manh.append(cur)
                    cur = ""
                if self._uoc_luong_dong(c, level, width_in) <= toi_da:
                    cur = c
                    continue
                for w in c.split():                      # một câu vẫn quá dài: tách theo từ
                    cand = f"{cur} {w}".strip()
                    if self._uoc_luong_dong(cand, level, width_in) <= toi_da:
                        cur = cand
                    else:
                        manh.append(cur)
                        cur = w
            if cur:
                manh.append(cur)
            for k, m in enumerate(manh):
                out.append((level, m, gach and k == 0))
        return out

    def _chia_trang(self, items, width_in):
        b = self.prof["body"]
        pages, cur, lines, n = [], [], 0, 0
        for it in items:
            need = self._uoc_luong_dong(it[1], it[0], width_in)
            if cur and (lines + need > b["max_lines"] or n + 1 > b["max_bullets"]):
                pages.append(cur)
                cur, lines, n = [], 0, 0
            cur.append(it)
            lines += need
            n += 1
        if cur:
            pages.append(cur)
        return pages or [[]]

    # ---------- hình ----------
    def _them_anh(self, s, blob, ext, x, y, w, h):
        """Đặt hình vừa khung (x, y, w, h), canh giữa. Hình hỏng thì báo cảnh báo, không dừng."""
        try:
            from PIL import Image
            iw, ih = Image.open(io.BytesIO(blob)).size
        except Exception:
            iw, ih = 4, 3
        scale = min(w / iw, h / ih)
        pw, ph = iw * scale, ih * scale
        try:
            s.shapes.add_picture(io.BytesIO(blob), Inches(x + (w - pw) / 2),
                                 Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))
        except Exception:
            self.warns.append(f"Slide “{s.shapes.title.text_frame.text}”: có hình không chèn được; "
                              "hãy chèn lại thủ công.")

    @staticmethod
    def _luoi_anh(n, x, y, w, h, gap):
        """Chia khung thành n ô cho n hình (1 ô; 2 ô cạnh nhau hoặc chồng nhau; 3-4 ô lưới 2x2)."""
        if n <= 1:
            return [(x, y, w, h)]
        if n == 2:
            if w / h > 1.5:
                cw = (w - gap) / 2
                return [(x, y, cw, h), (x + cw + gap, y, cw, h)]
            ch = (h - gap) / 2
            return [(x, y, w, ch), (x, y + ch + gap, w, ch)]
        cw, ch = (w - gap) / 2, (h - gap) / 2
        return [(x + (k % 2) * (cw + gap), y + (k // 2) * (ch + gap), cw, ch) for k in range(n)]

    def _dat_anh(self, s, anh, x, y, w, h):
        for (blob, ext), (cx, cy, cw, ch) in zip(anh, self._luoi_anh(len(anh), x, y, w, h,
                                                                       self.prof["image"]["gap_in"])):
            self._them_anh(s, blob, ext, cx, cy, cw, ch)

    # ---------- bảng ----------
    def _cao_hang(self, row, widths, size):
        """Ước lượng chiều cao một hàng (inch) theo số dòng chữ phải xuống."""
        tb = self.prof["table"]
        lines = 1
        for j, c in enumerate(row):
            if c in ("<", "^"):
                continue
            cpl = max(int((widths[j] - 2 * tb["cell_pad_in"]) * 72 / (size * 0.5)), 3)
            n = sum(max(1, math.ceil(len(x) / cpl)) for x in (c.split("\n") or [""]))
            lines = max(lines, n)
        return lines * size * 1.2 / 72 + tb["row_pad_in"]

    def _chia_bang(self, rows, widths, size):
        """Chia bảng thành các trang; mỗi trang lặp lại hàng đầu; không cắt ngang ô gộp dọc."""
        max_h = self.prof["body"]["height_in"]
        hh = self._cao_hang(rows[0], widths, size)
        pages, cur, h = [], [0], hh
        for i in range(1, len(rows)):
            rh = self._cao_hang(rows[i], widths, size)
            if len(cur) > 1 and h + rh > max_h and "^" not in rows[i]:
                pages.append(cur)
                cur, h = [0], hh
            cur.append(i)
            h += rh
        pages.append(cur)
        return pages

    def _them_bang(self, s, rows, size, widths):
        tb, b = self.prof["table"], self.prof["body"]
        nr, nc = len(rows), len(widths)
        heights = [self._cao_hang(r, widths, size) for r in rows]
        shp = s.shapes.add_table(nr, nc, Inches(b["left_in"]), Inches(b["top_in"]),
                                 Inches(sum(widths)), Inches(sum(heights)))
        tbl = shp.table
        for j, w in enumerate(widths):
            tbl.columns[j].width = Inches(w)
        for i, hgt in enumerate(heights):
            tbl.rows[i].height = Inches(hgt)
        for i, row in enumerate(rows):
            for j in range(nc):
                c = row[j] if j < len(row) else ""
                cell = tbl.cell(i, j)
                cell.margin_left = cell.margin_right = Inches(tb["cell_pad_in"])
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(self.c["primary"] if i == 0 else "FFFFFF" if i % 2 else self.c.get("zebra", "EEF2F8"))
                if c in ("<", "^"):
                    continue
                cell.text = ""
                tf = cell.text_frame
                tf.word_wrap = True
                for k, dong in enumerate(c.split("\n") or [""]):
                    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                    p.alignment = PP_ALIGN.CENTER if (i == 0 or (len(c) <= 4 and c.replace(".", "").isdigit())) \
                        else PP_ALIGN.LEFT
                    r = p.add_run()
                    r.text = dong
                    self._style_run(r, size, bold=(i == 0), color=self.c["on_primary"] if i == 0 else None)
        for i in range(nr):                      # gộp ô: "<" gộp trái, "^" gộp trên
            for j in range(nc):
                c = rows[i][j] if j < len(rows[i]) else ""
                if c in ("<", "^"):
                    continue
                w = 1
                while j + w < nc and j + w < len(rows[i]) and rows[i][j + w] == "<":
                    w += 1
                h = 1
                while i + h < nr and j < len(rows[i + h]) and rows[i + h][j] == "^":
                    h += 1
                if w > 1 or h > 1:
                    try:
                        tbl.cell(i, j).merge(tbl.cell(i + h - 1, j + w - 1))
                    except Exception:
                        pass

    def _dung_bang(self, rows, tieu, tao, ten_slide=""):
        """Dựng một bảng lên một hoặc nhiều slide (tự chọn cỡ chữ, tách trang). tieu(loai) -> tiêu đề."""
        tb, b = self.prof["table"], self.prof["body"]
        nc = max(len(r) for r in rows)
        rows = [r + [""] * (nc - len(r)) for r in rows]
        widths = [x / 25.4 for x in tinh_do_rong_cot(rows, b["width_in"] * 25.4, tb["min_col_in"] * 25.4)]
        chon = None
        for size in tb["font_sizes_pt"]:
            pages = self._chia_bang(rows, widths, size)
            if len(pages) == 1:
                chon = (size, pages)
                break
        if chon is None:
            size = tb["paged_font_pt"]
            chon = (size, self._chia_bang(rows, widths, size))
        size, pages = chon
        if len(pages) > 1:
            self.warns.append(f"Slide “{ten_slide}”: bảng có {len(rows)} dòng, đã tách thành "
                              f"{len(pages)} slide (lặp lại hàng tiêu đề).")
        for k, idx in enumerate(pages):
            s = self._tao_slide(tieu("bảng" if k == 0 else "tiếp"))
            self._them_bang(s, [rows[i] for i in idx], size, widths)
            tao.append(s)

    # ---------- một slide nội dung ----------
    def slide_noi_dung(self, sl: SlideIn, stt: int):
        b, im = self.prof["body"], self.prof["image"]
        anh = list(sl.images)
        sl.items = self._rut_y(sl)
        if self._la_the(sl):
            s = self._tao_slide(sl.title)
            self._ve_the(s, sl.items)
            if sl.notes:
                s.notes_slide.notes_text_frame.text = sl.notes
            return
        co_chu = bool(sl.items)
        w = b["width_with_image_in"] if (co_chu and anh) else b["width_in"]
        for it in sl.items:
            if len(it[1]) > b["long_item_chars"] and self.prof.get("standard", {}).get("chunking", {}).get("long_item", "warn") != "keep":
                self.warns.append(f"Slide “{sl.title}”: có đoạn dài {len(it[1])} ký tự, "
                                  "nên rút gọn thành ý chính.")
                break
        tao = []
        dau = [True]

        def tieu(loai):
            if dau[0]:
                dau[0] = False
                return sl.title
            return f"{sl.title} ({loai})"

        if co_chu:                                          # chữ (kèm tối đa per_slide hình ở bên phải)
            for k, items in enumerate(self._chia_trang(self._chia_doan_dai(sl.items, w), w)):
                s = self._tao_slide(tieu("tiếp"))
                self._dien_gach(self._body_box(s, w), items, w)
                if k == 0 and anh:
                    n = min(len(anh), im["per_slide"])
                    self._dat_anh(s, anh[:n], im["right_x_in"], b["top_in"], im["right_w_in"],
                                  b["height_in"] - 0.2)
                    anh = anh[n:]
                tao.append(s)
        for rows in sl.tables:                              # bảng
            if rows and rows[0]:
                self._dung_bang(rows, tieu, tao, sl.title)
        k = 0
        while anh:                                          # hình còn lại, mỗi slide tối đa per_slide hình
            n = min(len(anh), im["per_slide"])
            s = self._tao_slide(tieu("hình" if k == 0 else "tiếp"))
            self._dat_anh(s, anh[:n], b["left_in"], b["top_in"], b["width_in"], b["height_in"] - 0.2)
            anh = anh[n:]
            tao.append(s)
            k += 1
        if not tao:
            tao.append(self._tao_slide(tieu("tiếp")))
        if sl.notes and tao:
            tao[0].notes_slide.notes_text_frame.text = sl.notes

    def slide_muc_luc(self, tieu_de_list: list[str]):
        ag = self.prof["agenda"]
        s = self._tao_slide(ag["title"])
        tf = self._body_box(s, self.prof["body"]["width_in"])
        for i, t in enumerate(tieu_de_list):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(10)
            self._bullet(p, 0, "num")
            r = p.add_run()
            r.text = re.sub(r"^(?:\d+|[IVX]+)\s*[.)]\s*", "", t)   # đã có số tự động
            self._style_run(r, self.prof["body"]["size_pt"])


def dung_slide(slides: list[SlideIn], info: ThongTin, profile: str = KIEU_SLIDE[KIEU_MAC_DINH],
               tuy_chinh: dict | None = None, tieu_de_moi: dict | None = None):
    """tuy_chinh: ghi đè thông số profile (cùng cấu trúc yaml); tieu_de_moi: {số thứ tự slide (từ 1) hoặc tiêu đề cũ: tiêu đề mới}."""
    from ..profile import ap_dung_tuy_chinh
    prof = ap_dung_tuy_chinh(load_profile(profile), tuy_chinh)
    mx = prof.get("standard", {}).get("chunking", {}).get("max_chunks")
    if mx:
        prof["body"]["max_bullets"] = mx
    items = list(slides)
    for k, sl in enumerate(items, 1):
        moi = (tieu_de_moi or {}).get(str(k)) or (tieu_de_moi or {}).get(sl.title)
        if moi and moi.strip():
            sl.title = moi.strip()
    b = SlideBuilder(prof, info)
    ts = prof["title_slide"]

    # slide đầu chỉ có tiêu đề (+ tối đa 2 dòng phụ) = slide mở đầu
    ten_bai = info.ten_bai
    if items and len(items[0].items) <= 2 and not items[0].images and not items[0].tables:
        dau = items.pop(0)
        ten_bai = ten_bai or dau.title
        if not (info.hoc_phan or info.giang_vien or info.don_vi or info.nam):
            b.info.don_vi = " – ".join(t for _, t, _ in dau.items)
    ten_bai = ten_bai or ts.get("default_title", "BÀI GIẢNG")
    b.info.ten_bai = ten_bai
    if not (b.info.giang_vien and b.info.don_vi):
        b.warns.append("Chưa nhập đủ người trình bày/đơn vị cho slide mở đầu.")

    # slide cuối đã là lời cảm ơn/hỏi đáp: dùng làm slide kết, không tạo thêm một slide kết nữa
    ket_goc = None
    if prof.get("closing", {}).get("enabled") and items:
        cuoi = items[-1]
        if RE_CAM_ON.search(cuoi.title) and len(cuoi.items) <= 3 and not cuoi.images and not cuoi.tables:
            ket_goc = items.pop()

    # đưa slide "Tài liệu tham khảo" xuống cuối
    tltk = [s for s in items if RE_TLTK.search(s.title)]
    items = [s for s in items if s not in tltk] + tltk

    b.slide_tieu_de(ten_bai)
    ag = prof["agenda"]
    muc = [s.title for s in items if not RE_TLTK.search(s.title)]
    if ag["enabled"] and ag["min_items"] <= len(muc) <= ag["max_items"]:
        b.slide_muc_luc(muc)
    for i, sl in enumerate(items, 1):
        try:
            b.slide_noi_dung(sl, i)
        except Exception as e:             # một slide lỗi không làm hỏng cả bài
            b.warns.append(f"Slide “{sl.title}”: không dựng được ({type(e).__name__}); "
                           "hãy làm lại slide này thủ công.")
            try:
                b._tao_slide(sl.title)
            except Exception:
                pass
    if prof.get("closing", {}).get("enabled"):
        if ket_goc:
            b.slide_ket(ket_goc.title, [t for _, t, _ in ket_goc.items])
        else:
            b.slide_ket("", b._dong_thong_tin() if prof["closing"].get("show_contact") else None)
    buf = io.BytesIO()
    b.prs.save(buf)
    return buf.getvalue(), b.warns, len(b.prs.slides)
