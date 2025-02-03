from docx import Document

def delete_paragraph(paragraph):
    """删除指定段落"""
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

# 读取 Word 文档
doc = Document('technical_document_template.docx')

# **删除特定的一级标题及其后续内容，直到下一个一级标题**
delete_flag = False  # 标志是否要删除

for para in doc.paragraphs[:]:  # 复制列表避免遍历过程中修改
    if para.style.name == "Heading 1":  # 发现新的一级标题
        if para.text == "插入损耗特性":
            delete_flag = True  # 启动删除模式
        else:
            delete_flag = False  # 遇到新的一级标题，停止删除

    if delete_flag:
        delete_paragraph(para)  # 删除当前段落

# 保存修改后的文档
doc.save('example_modified.docx')
