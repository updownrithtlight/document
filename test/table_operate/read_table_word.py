# 获取所有表格
from docx import Document

# 读取 Word 文件
doc = Document("technical_document_template.docx")

# 遍历所有表格
for table_index, table in enumerate(doc.tables):
    print(f"📌 Table {table_index + 1} (共 {len(doc.tables)} 个表格)")

    # 遍历行
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            print(f"  📍 第 {row_index + 1} 行, 第 {col_index + 1} 列: {cell.text.strip()}")

    print("-" * 50)  # 分隔不同表格

from docx import Document


def replace_text_preserve_format(cell, replacements):
    """替换 Word 表格中的占位符，同时保持原格式"""
    for para in cell.paragraphs:
        for run in para.runs:
            for key, value in replacements.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)  # 替换内容，但不影响格式


def replace_text_in_table(doc, table_index, replacements):
    """在指定表格中批量替换占位符"""
    table = doc.tables[table_index]

    for row in table.rows:
        for cell in row.cells:
            replace_text_preserve_format(cell, replacements)


# 读取 Word 文件
another_doc = Document("technical_document_template.docx")

# 变量替换映射
replacements = {
    "{{project_model}}": "模型A",
    "{{project_name}}": "智能设备",
    "{{file_number}}": "编号-12345"
}

# 替换第 1 个表格（table_index = 0）
replace_text_in_table(another_doc, table_index=0, replacements=replacements)

# 保存修改后的 Word 文件
another_doc.save("modified_template.docx")


def delete_table_rows(doc, table_index, start_row, end_row):
    """
    删除 Word 指定表格中的行（通用方法）
    :param doc: Document 对象
    :param table_index: 需要操作的表格索引 (0-based)
    :param start_row: 删除的起始行索引 (0-based)
    :param end_row: 删除的结束行索引 (0-based)
    """
    if table_index >= len(doc.tables):
        raise ValueError(f"表格索引 {table_index} 超出范围，当前文档仅有 {len(doc.tables)} 个表格")

    table = doc.tables[table_index]
    total_rows = len(table.rows)

    # 处理超出行范围的情况
    if start_row >= total_rows or end_row >= total_rows:
        raise ValueError(f"表格 {table_index + 1} 仅有 {total_rows} 行，删除范围 ({start_row} - {end_row}) 超出范围")

    # 倒序删除，防止索引变化
    for row_index in range(end_row, start_row - 1, -1):
        table._element.remove(table.rows[row_index]._element)


# 加载 Word 文档
doc = Document("technical_document_template.docx")

# 删除【表2】的【第3行到第12行】（索引从 0 开始）
delete_table_rows(doc, table_index=2, start_row=2, end_row=11)

# 保存修改后的 Word 文件
doc.save("deleted_table_modified.docx")

