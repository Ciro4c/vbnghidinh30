import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from lxml import etree
from PIL import Image
def img(c):
    b=io.BytesIO(); Image.new("RGB",(400,300),c).save(b,"PNG"); b.seek(0); return b
prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
def tb(s,x,y,w,h,t,sz=20,bold=False):
    b=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); b.text_frame.text=t
    for r in b.text_frame.paragraphs[0].runs: r.font.size=Pt(sz); r.font.bold=bold
    return b
L=prs.slide_layouts[6]
# 1 sơ đồ
s=prs.slides.add_slide(L)
tb(s,1,0.4,11,0.8,"CÁC NGUYÊN NHÂN THƯỜNG DẪN ĐẾN SỰ CHÁY:",32)
tb(s,1.5,1.5,10,0.6,"ĐIỆN, GAS, HÀN XÌ, XĂNG DẦU, ĐỐT VÀNG MÃ, SỰ CỐ THIẾT BỊ...",18,True)
a=tb(s,1.5,2.6,2,1,"65%",40,True); tb(s,4.5,2.6,2,1,"35%",40,True); tb(s,7,2.6,3,1,"100%",40,True)
sh=tb(s,1.5,4.2,1,0.8,"SL",36); tb(s,1.5,5.3,1,0.8,"CL",36)
w=tb(s,5,4.6,4,1,"HQĐ 2500°C",36)
s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,Inches(2.8),Inches(4.4),Inches(1.5),Inches(0.4))
s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(2.8),Inches(5.4),Inches(4.8),Inches(5.4))
# bọc w trong AlternateContent
MC="http://schemas.openxmlformats.org/markup-compatibility/2006"
el=w._element; par=el.getparent()
ac=etree.SubElement(par,"{%s}AlternateContent"%MC,nsmap={"mc":MC}); ch=etree.SubElement(ac,"{%s}Choice"%MC); ch.set("Requires","a14")
par.remove(el); ch.append(el)
# 2 bảng
s=prs.slides.add_slide(L)
tb(s,1,0.3,11,1.2,"BẢNG II\nBÌNH CHỮA CHÁY CÁC LOẠI, THIẾT BỊ BÁO CHÁY",16,True)
t=s.shapes.add_table(3,3,Inches(1),Inches(2),Inches(10),Inches(2)).table
for i in range(3):
    for j in range(3): t.cell(i,j).text=f"{i}{j}"
# 3 ảnh + nhãn đè
s=prs.slides.add_slide(L)
tb(s,1,0.3,11,0.8,"Trường hợp bạn đã bị ngọn lửa bắt vào người thì ta phải làm gì?",24,True)
s.shapes.add_picture(img("blue"),Inches(1),Inches(1.8),Inches(5),Inches(4)); s.shapes.add_picture(img("green"),Inches(7),Inches(1.8),Inches(5),Inches(4))
tb(s,8,2.2,3,0.6,"Không chạy",24,True)
# 4 thống kê
s=prs.slides.add_slide(L)
b=tb(s,1,0.3,11,0.8,"I. THÔNG BÁO TÌNH HÌNH CHÁY, NỔ NĂM 2025",28,True)
x=tb(s,1,1.5,6,5,"Xảy ra 3060 vụ cháy",22)
tf=x.text_frame
for t_,lv in [("Làm chết 90 người",0),("Bị thương 93 người",0),("Thiệt hại khoảng 786,78 tỷ đồng",0),("Thống kê tháng 02 năm 2026",0),("Toàn quốc xảy ra: 254 vụ cháy",1),("Làm chết: 10 người",1),("Làm bị thương: 14 người",1),("Thiệt hại tài sản ước tính sơ bộ thành tiền khoảng: 30,11 tỷ",1)]:
    p=tf.add_paragraph(); p.text=t_; p.level=lv
s.shapes.add_picture(img("red"),Inches(8),Inches(1.5),Inches(4.5),Inches(4.5))
prs.save("/tmp/t/src.pptx")
