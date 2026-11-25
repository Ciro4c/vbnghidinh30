import docx, copy
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
d = docx.Document()
d.add_paragraph("TRƯỜNG ĐẠI HỌC ABC")
d.add_paragraph("ĐỀ CƯƠNG HỌC PHẦN").runs[0].bold = True
sdt = parse_xml(f'''<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>
<w:sdtContent>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Nội dung</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>LỜI GIỚI THIỆU</w:t></w:r><w:r><w:tab/><w:t>1</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>CHƯƠNG I: GIỚI THIỆU</w:t></w:r><w:r><w:tab/><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
</w:sdtContent></w:sdt>''')
d.element.body.insert(2, sdt)
d.add_paragraph("LỜI GIỚI THIỆU", style="Heading 1")
d.add_paragraph("Nội dung lời giới thiệu của giáo trình.")
d.add_paragraph("CHƯƠNG I: GIỚI THIỆU VỀ QUẢN TRỊ SẢN XUẤT", style="Heading 1")
d.add_paragraph("1.1 Thông tin chung", style="Heading 2")
d.add_paragraph("1.1.1 Giới thiệu tóm tắt", style="Heading 3")
d.add_paragraph("Nội dung chương một.")
d.add_paragraph("1.1.4.1 Giáo trình", style="Heading 4")
p = d.add_paragraph("Tích: ")
p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:sym w:font="Wingdings" w:char="F078"/></w:r>'))
p.add_run(" Cơ sở ngành; ")
p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:sym w:font="Wingdings" w:char="F0A8"/></w:r>'))
p.add_run(" Tự chọn")
r = p.add_run(" chữ nghiêng"); r.italic = True
d.add_paragraph("CHƯƠNG II: KẾT QUẢ", style="Heading 1")
d.add_paragraph("Nội dung chương hai.")
d.add_paragraph("PHỤ LỤC 1", style="Heading 1")
d.add_paragraph("Bảng phụ lục.")
d.save("out_toc.docx")
