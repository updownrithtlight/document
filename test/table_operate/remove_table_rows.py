from docx import Document


def find_row_index_by_content(table, content):
    """
    在 table 中查找包含指定 content 的行索引（0-based）。
    如果找不到，则返回 -1。
    """
    for i, row in enumerate(table.rows):
        # 合并该行所有单元格文本
        row_text = "".join(cell.text for cell in row.cells)
        if content in row_text:
            return i
    return -1


def delete_rows_in_range(table, start_row, end_row):
    """
    在给定的表格 table 中，删除从 start_row 到 end_row 行（含边界）。
    要求：start_row <= end_row。
    采用逆序删除，确保索引不混乱。
    """
    for row_idx in range(end_row, start_row - 1, -1):
        table._tbl.remove(table.rows[row_idx]._tr)


def process_missing_sections(
        doc_path,  # 输入Word文档路径
        table_index,  # 要操作的第几个表格（0-based）
        headings,  # 标题顺序列表
        missing_headings,  # 哪些标题是“没填写”的
        output_path  # 输出Word文档路径
):
    """
    若某标题在 missing_headings 中，则删除从“该标题”到“下一个标题”之前所有行。
    如果是最后一个标题，则一直删到表格末尾。
    """
    doc = Document(doc_path)
    table = doc.tables[table_index]

    # 1) 找到每个标题在表格中的行索引
    heading_row_map = {}
    for heading in headings:
        row_idx = find_row_index_by_content(table, heading)
        heading_row_map[heading] = row_idx  # 即使等于 -1，也先放进字典

    # 2) 为了防止多段连续删除导致的索引错乱，我们先把“要删除的标题”排序，
    #    按照它们在表格中的行索引从大到小排序，再依次删除
    #    注：如果有标题没找到(row=-1)，要特殊处理，示例里直接跳过
    def get_row_index(h):
        return heading_row_map.get(h, -1)

    # 从大到小排序 missing_headings
    missing_headings_sorted = sorted(
        missing_headings,
        key=get_row_index,
        reverse=True
    )

    # 3) 依次删除
    for i, heading in enumerate(missing_headings_sorted):
        start_row = heading_row_map.get(heading, -1)
        if start_row == -1:
            # 说明该标题在表格里根本就没找到
            # 根据实际需求处理，这里简单跳过
            continue

        # 获取“下一个标题”的行索引
        heading_index = headings.index(heading)
        next_index = heading_index + 1

        if next_index < len(headings):
            # 存在下一个标题
            next_heading = headings[next_index]
            next_heading_row = heading_row_map.get(next_heading, -1)
            # 如果也没找到，就视作要删除到末尾
            if next_heading_row == -1:
                end_row = len(table.rows) - 1
            else:
                end_row = next_heading_row - 1
        else:
            # 该标题已经是最后一个
            end_row = len(table.rows) - 1

        if end_row >= start_row:
            delete_rows_in_range(table, start_row, end_row)

    doc.save(output_path)
    print(f"处理完成，已保存到 {output_path}")


if __name__ == "__main__":
    # 需要先在你的程序中定义 headings 的顺序
    headings = [
        "电源部分",
        "信号部分",
        "电源输入特性",
        "电源输出特性",
        "特殊功能",
        "隔离特性"
    ]

    # 假设以下标题属于“没填写”的
    # 例如：用户没有输入“电源部分”和“电源输出特性”，
    # 则我们要删除从“电源部分”到“信号部分”，以及从“电源输出特性”到“特殊功能”（不含“特殊功能”）
    missing_headings = [
        "电源部分",
        "电源输出特性"
    ]

    process_missing_sections(
        doc_path="technical_document_template.docx",
        table_index=2,
        headings=headings,
        missing_headings=missing_headings,
        output_path="output.docx"
    )
