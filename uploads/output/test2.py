from docx import Document

def print_word_footers(template_path):
    """ 读取并打印 Word 文档中的所有页脚内容 """
    doc = Document(template_path)

    print("\n=== 📌 解析 Word 文档页脚内容 ===\n")
    for i, section in enumerate(doc.sections):
        print(f"\n🔹 Section {i + 1} 页脚内容:")
        footer = section.footer
        if footer:
            for j, para in enumerate(footer.paragraphs):
                print(f"  ➜ 段落 {j + 1}: {para.text.strip()}")
        else:
            print("  🚨 该页没有页脚！")

print_word_footers("../template/technical_document_template.docx")
