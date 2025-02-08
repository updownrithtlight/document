from docx import Document

class SpecWordTableProcessor:
    def __init__(self, file_path):
        """
        初始化工具类
        :param file_path: str, 输入的 Word 文档路径
        """
        self.file_path = file_path
        self.doc = Document(file_path)

    def process_table(self, selected_items, target_table_index=None, table_header=None):
        """
        处理指定的表格
        :param selected_items: list, 已选择项目的列表
        :param target_table_index: int, 指定表格的索引（从 0 开始），默认为 None
        :param table_header: str, 用于识别目标表格的表头文本，如果提供会优先通过表头匹配表格
        :return: None
        """
        # 提取已选择的项目名称
        selected_names = [item['name'] for item in selected_items]
        # Map for replacing True/False to symbols
        boolean_map = {True: '√', False: ''}
        # 定位目标表格
        target_table = None

        if table_header:
            # 如果指定了表头，通过表头识别表格
            for table in self.doc.tables:
                if table.rows[0].cells[0].text.strip() == table_header:
                    target_table = table
                    break
        elif target_table_index is not None:
            # 如果指定了表格索引，通过索引定位表格
            target_table = self.doc.tables[target_table_index]

        if not target_table:
            raise ValueError("未找到符合条件的表格，请检查参数 `table_header` 或 `target_table_index` 是否正确。")

        # 遍历目标表格并处理数据
        rows_to_delete = []  # 保存需要删除的行
        for row_idx, row in enumerate(target_table.rows[1:], start=1):  # 跳过表头
            cell_name = row.cells[0].text.strip()
            if cell_name not in selected_names:
                # 如果不在已选择项目中，标记为删除
                rows_to_delete.append(row)
            else:
                # 如果在已选择项目中，更新内容
                for item in selected_items:
                    if item['name'] == cell_name:
                        row.cells[1].text = boolean_map[item['pcb']]
                        row.cells[2].text = boolean_map[item['beforeSeal']]
                        row.cells[3].text = boolean_map[item['afterLabel']]
                        row.cells[4].text = item['samplePlan']

        # 删除不需要的行
        for row in rows_to_delete:
            target_table._element.remove(row._element)

    def save(self, output_path):
        """
        保存文档
        :param output_path: str, 输出的 Word 文档路径
        :return: None
        """
        self.doc.save(output_path)
