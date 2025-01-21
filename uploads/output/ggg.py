import zipfile
from lxml import etree


def replace_word_xml(docx_path, output_path, replacements):
    """ 直接修改 Word XML 进行占位符替换 """
    with zipfile.ZipFile(docx_path, 'r') as docx:
        new_docx = zipfile.ZipFile(output_path, 'w')

        for name in docx.namelist():
            content = docx.read(name)

            # **只处理包含 `document.xml` 和 `header` 的 XML 文件**
            if "document.xml" in name or "header" in name:
                xml_text = content.decode("utf-8")

                # **替换占位符**
                for placeholder, value in replacements.items():
                    xml_text = xml_text.replace(placeholder, value)

                new_docx.writestr(name, xml_text.encode("utf-8"))
            else:
                new_docx.writestr(name, content)

        new_docx.close()
    print(f"✅ 文档已生成: {output_path}")


# **占位符替换内容**
replacements = {
    "{{project_model}}": "MTLB32B-HNBJ",
    "{{file_number}}": "221226-ABC"
}

replace_word_xml("../template/technical_document_template.docx", "output.docx", replacements)
