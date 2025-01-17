import zipfile
from lxml import etree

def extract_footer_xml(docx_path):
    """ 解析 Word 页脚 XML，看看占位符是否存在 """
    with zipfile.ZipFile(docx_path, 'r') as docx:
        for name in docx.namelist():
            if "footer" in name:  # 找到页脚 XML 文件
                print(f"\n=== 🔍 解析页脚文件: {name} ===")
                xml_content = docx.read(name)
                root = etree.XML(xml_content)
                print(etree.tostring(root, pretty_print=True, encoding='utf-8').decode('utf-8'))

extract_footer_xml("technical_document_template.docx")
