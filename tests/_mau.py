# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Dựng file Word/PowerPoint thô có bảng, hình để kiểm thử."""
import io

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Mm
from PIL import Image, ImageDraw


def png(w=300, h=200, mau=(30, 90, 200)):
    im = Image.new("RGB", (w, h), mau)
    ImageDraw.Draw(im).line((0, 0, w, h), fill=(255, 255, 255), width=4)
    b = io.BytesIO()
    im.save(b, "PNG")
    b.seek(0)
    return b


def docx_co_bang_hinh() -> bytes:
    d = Document()
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "UBND TỈNH ABC\nSỞ GIÁO DỤC VÀ ĐÀO TẠO"
    t.cell(0, 1).text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"
    t.cell(1, 0).text = "Số: 7/SGDĐT-VP\nV/v hướng dẫn sinh hoạt chuyên môn"
    t.cell(1, 1).text = "Cần Thơ, ngày 5 tháng 1 năm 2026"
    d.add_paragraph("Kính gửi: Các trường THPT")
    d.add_paragraph("Thực hiện kế hoạch, Sở hướng dẫn như sau:")
    d.add_paragraph("1. Bảng số liệu").runs[0].bold = True
    bt = d.add_table(rows=5, cols=4)
    bt.style = "Table Grid"
    for j, h in enumerate(["STT", "Nội dung công việc", "Đơn vị", "Ghi chú"]):
        c = bt.cell(0, j)
        c.text = ""
        c.paragraphs[0].add_run(h).bold = True
    for i in range(1, 5):
        bt.cell(i, 0).text = str(i)
        bt.cell(i, 1).text = f"Công việc số {i}"
        bt.cell(i, 2).text = "Phòng A\nPhòng B"
    bt.cell(1, 3).merge(bt.cell(2, 3))
    bt.cell(1, 3).text = "Gộp dọc"
    bt.cell(4, 1).merge(bt.cell(4, 2))
    bt.cell(4, 1).text = "Gộp ngang"
    d.add_paragraph("2. Hình minh họa").runs[0].bold = True
    p = d.add_paragraph()
    p.add_run().add_picture(png(900, 500), width=Mm(120))
    d.add_paragraph("Đoạn có hộp văn bản:")
    p = d.paragraphs[-1]
    p._p.append(parse_xml(
        '<w:r %s><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>'
        '<w:p><w:r><w:t>Chữ trong hộp</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>'
        % nsdecls("w")))
    for s in ["Nơi nhận:", "- Như trên;", "- Lưu: VT.", "KT. GIÁM ĐỐC", "PHÓ GIÁM ĐỐC", "(Đã ký)", "NGUYỄN VĂN A"]:
        d.add_paragraph(s)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


def pptx_co_bang_hinh() -> bytes:
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "BÀI TEST"
    s.placeholders[1].text = "Khoa A"
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Bảng dài"
    t = s.shapes.add_table(30, 4, Inches(0.5), Inches(1.5), Inches(12), Inches(5)).table
    for r in range(30):
        for c in range(4):
            t.cell(r, c).text = f"R{r}C{c}"
    t.cell(0, 0).merge(t.cell(0, 1))
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Năm ảnh"
    for i in range(5):
        s.shapes.add_picture(png(4000, 3000, (40 * i, 100, 200)), Inches(0.5 + i), Inches(2), Inches(1))
    b = io.BytesIO()
    prs.save(b)
    return b.getvalue()
