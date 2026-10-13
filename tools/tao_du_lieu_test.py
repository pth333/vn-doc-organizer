"""Tao thu muc test voi cac truong hop dac biet de kiem tra tool."""
import os
import zipfile

OUT = r"H:\tool_files\_test_dau_vao"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def tao_docx(path, lines):
    body = "".join("<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" % l for l in lines)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s</w:body></w:document>' % body)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc)


# 1. .docx co day du thong tin
tao_docx(os.path.join(OUT, "sach_moi_xuat_ban.docx"), [
    "Nguyễn Văn A",
    "Kinh tế Việt Nam thời kỳ đổi mới",
    "nhà xuất bản chính trị quốc gia",
    "Hà Nội - 2015",
    "Chủ biên: PGS. TS. Nguyễn Văn A",
    "Lời nhà xuất bản",
    "Cuốn sách được xuất bản năm 2015 nhân dịp kỷ niệm.",
])

# 2. .docx tai ban, khong co nam ro rang -> phai vao _KHONG_XAC_DINH_NAM
tao_docx(os.path.join(OUT, "tai_lieu_khong_co_nam.docx"), [
    "Tài liệu hướng dẫn nghiệp vụ",
    "Nội dung không ghi năm ban hành ở đâu cả.",
    "Tái bản lần thứ 3, có sửa chữa bổ sung.",
])

# 3. .docx co nam nhung khong co ngay
tao_docx(os.path.join(OUT, "sach_chi_co_nam_2010.docx"), [
    "Một số vấn đề lý luận",
    "Hà Nội - 2010",
    "nhà xuất bản chính trị quốc gia",
])

# 4. PDF hong (khong phai PDF that)
with open(os.path.join(OUT, "file_pdf_hong.pdf"), "wb") as fh:
    fh.write(b"%PDF-1.4\n day khong phai file pdf hop le \n%%EOF\n")

# 5. File khac dinh dang
with open(os.path.join(OUT, "ghi_chu.txt"), "w", encoding="utf-8") as fh:
    fh.write("file txt - khong duoc ho tro\n")

print("Da tao:", os.listdir(OUT))
