from docx import Document


class WordTableProcessor:
    def __init__(self, doc_path, table_index=0):
        """
        初始化处理器，加载Word文档并锁定目标表格。

        :param doc_path:    Word文档路径
        :param table_index: 要操作的表格索引（从0开始）
        """
        self.doc = Document(doc_path)
        self.table = self.doc.tables[table_index]
        self.doc_path = doc_path
        self.table_index = table_index

    def _find_row_index_by_content(self, content):
        """
        在表格中查找包含指定 content 的行索引（0-based）。
        如果找不到，则返回 -1。
        """
        for i, row in enumerate(self.table.rows):
            row_text = "".join(cell.text for cell in row.cells)
            if content in row_text:
                return i
        return -1

    def _delete_rows_in_range(self, start_row, end_row):
        """
        在表格中删除从 start_row 到 end_row 行（含两端）。
        要求：start_row <= end_row。
        采用逆序删除，确保索引不混乱。
        """
        for row_idx in range(end_row, start_row - 1, -1):
            self.table._tbl.remove(self.table.rows[row_idx]._tr)

    def process_missing_sections(self, headings, missing_headings, output_path):
        """
        若某标题在 missing_headings 中，则删除从“该标题”到“下一个标题”之前所有行。
        如果是最后一个标题，则一直删到表格末尾。

        :param headings:         所有标题的顺序列表
        :param missing_headings: 未填写的标题列表
        :param output_path:      处理完成后存储的新文档路径
        """
        # 1) 找到每个标题在表格中的行索引
        heading_row_map = {}
        for heading in headings:
            row_idx = self._find_row_index_by_content(heading)
            heading_row_map[heading] = row_idx

        # 2) 为了防止多段连续删除导致的索引错乱，先把“要删除的标题”按照
        #    它们在表格中的行索引从大到小排序，再依次删除
        def get_row_index(h):
            return heading_row_map.get(h, -1)

        missing_headings_sorted = sorted(missing_headings, key=get_row_index, reverse=True)

        # 3) 逐个标题对应删除区间
        for heading in missing_headings_sorted:
            start_row = heading_row_map.get(heading, -1)
            if start_row == -1:
                # 标题没在表格中找到，跳过
                continue

            # 找到此标题在 headings 中的位置，确定下一个标题
            heading_index = headings.index(heading)
            next_index = heading_index + 1

            if next_index < len(headings):
                # 存在下一个标题
                next_heading = headings[next_index]
                next_heading_row = heading_row_map.get(next_heading, -1)

                # 如果下一个标题在表格中没找到，就删除到表格末尾
                if next_heading_row == -1:
                    end_row = len(self.table.rows) - 1
                else:
                    end_row = next_heading_row - 1
            else:
                # 该标题已经是最后一个了
                end_row = len(self.table.rows) - 1

            if end_row >= start_row:
                self._delete_rows_in_range(start_row, end_row)

        # 4) 保存结果
        self.doc.save(output_path)
        print(f"处理完成，已保存到 {output_path}")

