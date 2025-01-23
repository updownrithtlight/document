from docx.oxml import OxmlElement
from docx import Document
doc = Document('technical_document_template.docx')


def delete_paragraph(paragraph):
    """删除指定段落"""
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None


# 删除特定标题
for para in doc.paragraphs:
    if para.style.name.startswith("Heading") and para.text == "修改后的一级标题":
        delete_paragraph(para)

# 删除特定段落
for para in doc.paragraphs:
    if para.text == "这是修改后的第一个段落。":
        delete_paragraph(para)

doc.save('example_modified.docx')
