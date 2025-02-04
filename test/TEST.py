import os
import win32com.client as win32
from docx import Document

def delete_section_by_title(doc, target_title, heading_level="Heading 1"):
    """
    删除 Word 文档中，从指定标题开始直到下一个同级标题出现之前的所有段落。
    注意：该方法依赖 python‑docx 内部对象，未来版本可能需要调整。
    """
    def delete_paragraph(paragraph):
        element = paragraph._element
        element.getparent().remove(element)
        # 将内部对象置空，确保段落被删除
        paragraph._p = paragraph._element = None

    delete_flag = False
    for para in doc.paragraphs[:]:
        if para.style.name == heading_level:
            if para.text.strip() == target_title:
                delete_flag = True
            else:
                delete_flag = False
        if delete_flag:
            delete_paragraph(para)

def update_toc_via_word(doc_path):
    """
    利用 Microsoft Word COM 自动化接口更新文档目录（TOC）。
    打开指定文档，调用目录的 Update 方法后保存并关闭文档。
    """
    abs_path = os.path.abspath(doc_path)
    print("启动 Word 应用程序...")
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 后台运行

    try:
        print("打开文档：", abs_path)
        doc = word.Documents.Open(abs_path)
    except Exception as e:
        print("❌ 打开文档失败：", e)
        word.Quit()
        return False

    try:
        if doc.TablesOfContents.Count > 0:
            toc = doc.TablesOfContents(1)
            print("更新目录 TOC ...")
            toc.Update()
        else:
            print("文档中没有目录 TOC")
        print("保存文档...")
        doc.Save()
    except Exception as e:
        print("❌ 更新目录失败：", e)
        doc.Close(False)
        word.Quit()
        return False

    try:
        doc.Close(False)
        print("关闭文档成功。")
    except Exception as e:
        print("关闭文档时出现异常：", e)
    finally:
        word.Quit()
        print("Word 应用程序已退出。")
    return True

def main():
    # 模板文档路径（请根据实际情况修改）
    template_path = r"C:\noneSystem\bj\document\test\technical_document_template.docx"
    if not os.path.exists(template_path):
        print("❌ 模板文件不存在：", template_path)
        return

    # 生成修改后的文档路径
    new_doc_path = template_path.replace(".docx", "_modified.docx")

    print("📌 正在加载模板文档并删除指定章节...")
    try:
        doc = Document(template_path)
    except Exception as e:
        print("❌ 加载文档失败：", e)
        return

    print("📌 正在删除标题 '插入损耗特性' 及其后续内容...")
    delete_section_by_title(doc, "插入损耗特性")
    print("📌 正在删除标题 '产品功能' 及其后续内容...")
    delete_section_by_title(doc, "产品功能")

    try:
        doc.save(new_doc_path)
        print(f"✅ 文档修改成功，已保存到：{new_doc_path}")
    except Exception as e:
        print("❌ 保存文档失败：", e)
        return

    print("📌 正在调用 Word COM 自动化接口更新目录 TOC ...")
    if update_toc_via_word(new_doc_path):
        print("✅ TOC 更新成功")
    else:
        print("❌ TOC 更新失败")

if __name__ == "__main__":
    main()
