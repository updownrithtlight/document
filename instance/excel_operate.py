import openpyxl
from openpyxl.utils import get_column_letter

def fill_excel_template(template_path, output_path, product_info, data):
    # Load the workbook and select the active worksheet
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active

    # Fill fixed product info in specific cells
    def set_cell_value(cell, value):
        try:
            if not isinstance(sheet[cell], openpyxl.cell.cell.MergedCell):
                sheet[cell].value = value
        except AttributeError:
            pass  # Handle merged cell error gracefully

    set_cell_value("C2", product_info.get("产品名称", ""))
    set_cell_value("C3", product_info.get("产品型号", ""))
    set_cell_value("C4", product_info.get("成品规格", ""))
    set_cell_value("H1", product_info.get("文件编号", ""))
    set_cell_value("H3", product_info.get("产品等级", ""))
    set_cell_value("H4", product_info.get("产品编号", ""))

    # Insert data starting from row 6 (after the fixed headers)
    for idx, row_data in enumerate(data, start=6):
        try:
            sheet[f"A{idx}"].value = row_data.get("序号", "")
            sheet[f"B{idx}"].value = row_data.get("物品编码", "")
            sheet[f"C{idx}"].value = row_data.get("元器件名称", "")
            sheet[f"D{idx}"].value = row_data.get("型号", "")
            sheet[f"E{idx}"].value = row_data.get("单位", "")
            sheet[f"F{idx}"].value = row_data.get("数量", "")
            if not isinstance(sheet[f"H{idx}"], openpyxl.cell.cell.MergedCell):
                sheet[f"H{idx}"].value = row_data.get("参数/备注", "")
        except AttributeError:
            pass  # Handle merged cell error gracefully

    # Save the updated workbook
    wb.save(output_path)

# Example usage:
input_file = "MTLB32B-HNBJ-8A-KRW BOM物料表【优化】 231020.xlsx"
output_file = "Filled_BOM_Template.xlsx"

# Sample product info to be filled
product_info = {
    "产品名称": "电源滤波组件",
    "产品型号": "MTLB32B-HNBJ-8A-KRW",
    "成品规格": "275VAC 50Hz 8A 单相",
    "文件编号": "0-0WL",
    "产品等级": "J",
    "产品编号": "22199"
}

# Sample data to be filled
sample_data = [
    {"序号": 1, "物品编码": "001", "元器件名称": "电容器", "型号": "C1", "单位": "个", "数量": 10, "参数/备注": "说明1"},
    {"序号": 2, "物品编码": "002", "元器件名称": "电阻器", "型号": "R1", "单位": "个", "数量": 20, "参数/备注": "说明2"},
    {"序号": 3, "物品编码": "003", "元器件名称": "二极管", "型号": "D1", "单位": "个", "数量": 5, "参数/备注": "说明3"}
]

# Fill the template with product info and sample data
fill_excel_template(input_file, output_file, product_info, sample_data)
