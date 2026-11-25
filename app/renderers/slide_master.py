# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Định dạng lại Slide Master và các Layout của file slide theo profile.

Mẫu mặc định của python-pptx là khổ 4:3, chữ Calibri/Times, tiêu đề căn giữa; nếu không chỉnh thì slide mới
người dùng thêm trong PowerPoint (Home > New Slide) lệch hẳn so với slide đã dựng. Hàm `ap_master` làm cho
master, layout và theme (phông, màu) khớp profile: khổ 16:9, vùng tiêu đề/nội dung/chân trang theo %, phông,
cỡ chữ, màu, và đưa các nét trang trí (khối nhấn, đường kẻ dưới tiêu đề, đường kẻ chân trang) vào layout
\"Title Only\" để slide mới có sẵn."""
from __future__ import annotations

import re

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

LAYOUT_GIU = ("Title Slide", "Title and Content", "Title Only", "Blank")
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def _emu(v: float) -> int:
    return int(Inches(v))


def _theme(prs, prof):
    c, font = prof["colors"], prof["font"]
    try:
        part = prs.slide_master.part.part_related_by(RT.THEME)
        xml = part.blob.decode("utf-8")
    except Exception:
        return
    for tag in ("majorFont", "minorFont"):
        xml = re.sub(r"(<a:%s>\s*<a:latin typeface=\")[^\"]*(\")" % tag, r"\g<1>%s\g<2>" % font, xml)
    for tag, key in (("accent1", "accent"), ("accent2", "highlight"), ("dk2", "primary")):
        if key in c:
            xml = re.sub(r"(<a:%s>\s*<a:srgbClr val=\")[0-9A-Fa-f]{6}(\")" % tag, r"\g<1>%s\g<2>" % c[key], xml)
    for tag, hexc in (prof.get("theme_colors") or {}).items():          # bảng màu theme theo tiêu chuẩn (dk1, lt1, accent1...)
        xml = re.sub(r"<a:%s>.*?</a:%s>" % (tag, tag),
                     '<a:%s><a:srgbClr val="%s"/></a:%s>' % (tag, hexc, tag), xml, count=1, flags=re.S)
    part._blob = xml.encode("utf-8")


def _fill(parent, hexc, truoc=None):
    for old in parent.findall(qn("a:solidFill")):
        parent.remove(old)
    sf = etree.Element(qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", hexc)
    parent.insert(0, sf)


def _kieu_chu(master, prof):
    c, ti, bo = prof["colors"], prof["title"], prof["body"]
    ts = master._element.find(qn("p:txStyles"))
    if ts is None:
        return
    lv = ts.find("p:titleStyle/a:lvl1pPr", NS)
    if lv is not None:
        lv.set("algn", "l")
        d = lv.find(qn("a:defRPr"))
        if d is not None:
            d.set("sz", str(int(ti["size_pt"] * 100)))
            d.set("b", "1")
            _fill(d, c["primary"])
    for i in range(1, 10):
        lv = ts.find("p:bodyStyle/a:lvl%dpPr" % i, NS)
        if lv is None:
            continue
        d = lv.find(qn("a:defRPr"))
        if d is not None:
            d.set("sz", str(int((bo["size_pt"] if i == 1 else bo["sub_size_pt"]) * 100)))
            _fill(d, c["text"])


def _dat_vi_tri(sp, x, y, w, h):
    spPr = sp.find(qn("p:spPr"))
    xf = spPr.find(qn("a:xfrm"))
    if xf is None:
        xf = etree.Element(qn("a:xfrm"))
        spPr.insert(0, xf)
        etree.SubElement(xf, qn("a:off"))
        etree.SubElement(xf, qn("a:ext"))
    xf.find(qn("a:off")).set("x", str(_emu(x)))
    xf.find(qn("a:off")).set("y", str(_emu(y)))
    xf.find(qn("a:ext")).set("cx", str(_emu(w)))
    xf.find(qn("a:ext")).set("cy", str(_emu(h)))


def _bo_xfrm(sp):
    spPr = sp.find(qn("p:spPr"))
    if spPr is not None:
        for xf in spPr.findall(qn("a:xfrm")):
            spPr.remove(xf)


def _loai_ph(sp):
    ph = sp.find(".//p:nvPr/p:ph", NS)
    return None if ph is None else (ph.get("type") or "obj")


def _hinh_trang_tri(layout, prof, vung_chan):
    """Khối nhấn, đường kẻ dưới tiêu đề, đường kẻ chân trang -> layout Title Only."""
    c, tp, hd, ft = prof["colors"], prof["title"], prof.get("header", {}), prof["footer"]
    W = prof["slide"]["width_in"]
    tree = layout._element.find(qn("p:cSld")).find(qn("p:spTree"))
    nid = [900]

    def rect(ten, x, y, w, h, mau):
        nid[0] += 1
        xml = ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
               '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
               '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
               % (NS["p"], NS["a"], nid[0], ten, _emu(x), _emu(y), _emu(w), _emu(h), mau))
        tree.append(etree.fromstring(xml))

    def duong(ten, x1, x2, y, mau, pt):
        nid[0] += 1
        xml = ('<p:cxnSp xmlns:p="%s" xmlns:a="%s"><p:nvCxnSpPr><p:cNvPr id="%d" name="%s"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr>'
               '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
               '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr></p:cxnSp>'
               % (NS["p"], NS["a"], nid[0], ten, _emu(x1), _emu(y), _emu(x2 - x1), int(pt * 12700), mau))
        tree.append(etree.fromstring(xml))

    if hd.get("top_bar", True):
        rect("Thanh trên", 0, 0, W, 0.12, c["primary"])
    if hd.get("marker"):
        rect("Khối nhấn", tp["left_in"] - 0.3, tp["top_in"] + 0.17, 0.09, tp["height_in"] - 0.34, c["accent"])
    kieu = hd.get("underline", "dai")
    if kieu != "khong":
        y = tp["top_in"] + tp["height_in"] + 0.05
        mau = c["accent"] if kieu == "dai" else c.get("border", c["accent"])
        duong("Đường kẻ tiêu đề", tp["left_in"], W - tp["left_in"], y, mau, hd.get("underline_pt", 2))
    if ft.get("line"):
        xl = tp["left_in"] if (vung_chan or prof.get("master")) else 0.6
        yl = (vung_chan + 0.1) if vung_chan else ft.get("line_y_in", prof["slide"]["height_in"] - 0.62)
        duong("Đường kẻ chân trang", xl, W - xl, yl, c.get("border", c["muted"]), 0.75)


def ap_master(prs, prof, vung_chan=None) -> dict:
    """Trả về {'tieu_de': layout Title Slide, 'noi_dung': layout Title Only}."""
    c = prof["colors"]
    W, H = prof["slide"]["width_in"], prof["slide"]["height_in"]
    tp, bo, ft = prof["title"], prof["body"], prof["footer"]
    master = prs.slide_master
    _theme(prs, prof)
    _kieu_chu(master, prof)
    try:
        master.background.fill.solid()
        master.background.fill.fore_color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor.from_string(c["bg"])
    except Exception:
        pass

    xl = tp["left_in"]
    yl = (vung_chan + 0.1) if vung_chan else prof["footer"].get("line_y_in", H - 0.62)
    mt = {k: tuple(v) for k, v in (prof.get("master") or {}).items()}      # toạ độ khung chân trang theo tiêu chuẩn (inch)
    for sp in master.placeholders:
        t = _loai_ph(sp._element)
        if t in ("title", "ctrTitle"):
            _dat_vi_tri(sp._element, tp["left_in"], tp["top_in"], tp["width_in"], tp["height_in"])
        elif t in ("body", "obj"):
            _dat_vi_tri(sp._element, bo["left_in"], bo["top_in"], bo["width_in"], bo["height_in"])
        elif t == "dt":
            _dat_vi_tri(sp._element, *mt.get("date", (xl + 4.0, yl + 0.07, 4.0, 0.4)))
        elif t == "ftr":
            _dat_vi_tri(sp._element, *mt.get("footer", (xl, yl + 0.07, 5.0, 0.4)))
        elif t == "sldNum":
            _dat_vi_tri(sp._element, *mt.get("slide_number", (W - xl - 2.0, yl + 0.07, 2.0, 0.4)))

    # chỉ giữ vài layout hay dùng
    for lay in list(prs.slide_layouts):
        if lay.name not in LAYOUT_GIU:
            try:
                prs.slide_layouts.remove(lay)
            except Exception:
                pass
    kq = {}
    for lay in prs.slide_layouts:
        for sp in lay.placeholders:
            t = _loai_ph(sp._element)
            if lay.name == "Title Slide":
                if t == "ctrTitle":
                    _dat_vi_tri(sp._element, 1.0, 2.0, W - 2.0, 1.9)
                elif t == "subTitle":
                    _dat_vi_tri(sp._element, 1.0, 4.1, W - 2.0, 1.3)
            elif t in ("title", "body", "obj", "dt", "ftr", "sldNum"):
                _bo_xfrm(sp._element)                      # kế thừa vị trí từ master
        if lay.name == "Title Slide":
            kq["tieu_de"] = lay
        elif lay.name == "Title Only":
            kq["noi_dung"] = lay
            _hinh_trang_tri(lay, prof, vung_chan)
    return kq
