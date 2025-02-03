from docx import Document


def delete_section_by_title(doc, target_title, heading_level="Heading 1"):
    """
    删除 Word 文档中特定标题及其后续内容，直到下一个相同级别的标题。

    :param doc: Word 文档对象
    :param target_title: 要删除的标题文本
    :param heading_level: 要匹配的标题级别（默认 "Heading 1"）
    """

    def delete_paragraph(paragraph):
        """删除指定段落"""
        p = paragraph._element
        p.getparent().remove(p)
        paragraph._p = paragraph._element = None

    delete_flag = False  # 标志是否删除

    for para in doc.paragraphs[:]:  # 复制列表，避免遍历时修改
        if para.style.name == heading_level:  # 遇到新的相同级别标题
            if para.text.strip() == target_title:  # 标题匹配，开始删除
                delete_flag = True
            else:
                delete_flag = False  # 遇到下一个标题，停止删除

        if delete_flag:
            delete_paragraph(para)  # 删除当前段落


# **示例调用**
doc = Document('technical_document_template.docx')

# **删除特定标题及其内容**
delete_section_by_title(doc, "插入损耗特性")  # 仅删除 "插入损耗特性" 标题及其内容

# **保存修改后的文档**
doc.save('example_modified.docx')
