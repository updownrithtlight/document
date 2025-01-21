# import zipfile
# from lxml import etree
#
# def extract_word_xml(docx_path):
#     """ 解析 Word 文档的所有 XML 文件，查找 {{project_model}} """
#     with zipfile.ZipFile(docx_path, 'r') as docx:
#         for name in docx.namelist():
#             if name.endswith(".xml"):
#                 xml_content = docx.read(name).decode("utf-8")
#                 if "{{project_model}}" in xml_content:
#                     print(f"\n✅ 发现占位符 `{{project_model}}` 在 {name} 中！")
#                 if "{{file_number}}" in xml_content:
#                     print(f"\n✅ 发现占位符 `{{file_number}}` 在 {name} 中！")
#
# extract_word_xml("technical_document_template.docx")
