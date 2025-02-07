import json
from datetime import datetime
from flask import jsonify, send_file, request
from flask_jwt_extended import jwt_required
from urllib.parse import quote
from app import app
from app.utils.docx_processor import DocxProcessor
from app.utils.remove_image import process_section_by_marker
from app.utils.table_operation import add_row_with_auto_serial
from app.controllers.project_field_controller import get_list_by_project_id, get_fields_by_project_id_parent_id
from app.controllers.project_feature_controller import get_features
from app.controllers.project_important_notes_controller import get_important_notes
from app.controllers.field_definition_controller import get_fields_by_code, get_fields_h2_by_code
from app.exceptions.exceptions import CustomAPIException
from app.models.models import Project
import os
import zipfile
import shutil
from lxml import etree
from docx import Document

from app.utils.word_table_processor import WordTableProcessor
from app.utils.word_toc_tool import WordTocTool

# **模板文件路径**
PRODUCT_SPECIFICATION_TEMPLATE_PATH = os.path.join(app.config['TEMPLATE_FOLDER'], "product_specification.docx")


@jwt_required()
def generate_product_spec(project_id):
    """
    生成产品规范 Word 文档
    """
    try:
        project = Project.query.get(project_id)
        if not project:
            return jsonify({"error": "项目不存在"}), 404

        project_field_list = get_list_by_project_id(project_id)
        today = datetime.today()

        # 生成格式化的日期字符串
        formatted_date = today.strftime("%y%m%d")

        print(formatted_date)

        # **生成文件路径**
        output_file_name = f"{project.project_model}产品规范 {formatted_date}.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file_name)

        # **填充 Word 模板**
        fill_placeholder_template(PRODUCT_SPECIFICATION_TEMPLATE_PATH, output_path, project, project_field_list)
        # **更新目录**
        WordTocTool.update_toc_via_word(output_path)
        # **URL 编码文件名，避免中文乱码**
        encoded_file_name = quote(output_file_name)

        # **返回文件**
        response = send_file(
            output_path,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_file_name}"
        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500


def check_note_id_8(important_notes):
    """
    important_notes: 期望是一个list，每个元素都是dict
    """
    for item in important_notes.get("important_notes"):
        if not isinstance(item, dict):
            # item 不是dict，可能是字符串或其他类型，做一下容错或抛异常
            continue
        if item.get("note_id") == 8:
            return False
    return True



def demo_missing_headings(headings, id_to_name, table_part_ids):
    """
    根据给定的标题顺序(headings)、ID->标题字典(id_to_name)，以及用户已填写的ID列表(table_part_ids)，
    计算并返回“未填写”的标题列表。

    :param headings:       list[str]  所有标题（有固定顺序）
    :param id_to_name:     dict[int, str]  映射：ID -> 对应标题
    :param table_part_ids: list[int]  数据库查询得到的已填写标题的ID列表

    :return missing_headings: list[str]  未填写的标题名称列表
    """
    # 1) 将“已填写”的 ID 转成“已填写的标题”列表
    #    注：如果有ID在 id_to_name 中未找到，也可视情况处理，这里简单过滤
    found_headings = []
    for _id in table_part_ids:
        if _id in id_to_name:
            found_headings.append(id_to_name[_id])
        else:
            # 如果出现异常ID，不在字典中，可选择跳过或报错
            pass

    # 2) 用集合操作计算“所有标题”与“已填写标题”的差集
    missing_set = set(headings) - set(found_headings)
    # 再转回列表；如果想保持标题在 headings 中的原始顺序，可做个排序
    missing_headings = [h for h in headings if h in missing_set]

    return missing_headings


def build_cleaned_dict(filtered_part_2):
    cleaned_dict = {}
    skip_codes = {"fuse", "conductive_pad", "factory_test_report", "test_report"}

    for item in filtered_part_2:
        code = item.get("code")
        if code is None or code in skip_codes:
            continue  # 跳过 code 为空的情况

        # 生成字典的键，形如 "{{manufacturing_process}}"
        dict_key = f"{{{{{code}}}}}"

        if code == "manufacturing_process":
            # 1. 取得 custom_value
            raw_value = item.get("custom_value", "N/A")

            # 2. 解析 JSON，并用 "、" 连接
            try:
                data_list = json.loads(raw_value) if raw_value not in ["N/A", None, ""] else []
            except json.JSONDecodeError:
                data_list = []

            formatted_str = "、".join(data_list) if data_list else "N/A"

            # 3. 特殊处理后的值赋给 cleaned_dict
            cleaned_dict[dict_key] = formatted_str
        else:
            # 直接使用原 custom_value
            cleaned_dict[dict_key] = item.get("custom_value", "")

    return cleaned_dict

def build_placeholders(records):
    """
    :param records: 列表，形如:
        [
          {
            'code': 'PCV',
            'min_value': '1',
            'typical_value': '1',
            'max_value': '1',
            'unit': None,
            'description': '11',
            ...
          },
          ...
        ]
    :return: 一个字典，如:
        {
          '{{PCV1}}': '1',
          '{{PCV2}}': '1',
          '{{PCV3}}': '1',
          '{{PCV4}}': '--',
          '{{PCV5}}': '11',
          '{{PCC1}}': '1',
          ...
        }
    """
    result = {}
    for item in records:
        # 1. 获取 code
        code = item.get('code')
        if not code:
            # 如果没有 code，跳过或根据需求自行处理
            continue

        # 2. 逐个生成 5 个键 -> 值
        # 如果字段值为空或 None，就使用 "--"
        min_val = item.get('min_value') or "--"
        typ_val = item.get('typical_value') or "--"
        max_val = item.get('max_value') or "--"
        unit_val = item.get('unit') or "--"
        desc_val = item.get('description') or "--"

        result[f"{{{{{code}1}}}}"] = min_val
        result[f"{{{{{code}2}}}}"] = typ_val
        result[f"{{{{{code}3}}}}"] = max_val
        result[f"{{{{{code}4}}}}"] = unit_val
        result[f"{{{{{code}5}}}}"] = desc_val

    return result



def filter_missing_field_names(baseline_data, input_data):
    """
    过滤出在基准数据中存在，但在输入数据中没有出现的字段名称。

    :param baseline_data: 字典，包含基准数据，格式为 {code: field_data}
    :param input_data: 字典，包含要检查的数据，格式应与 baseline_data 相同
    :return: 列表，包含那些在输入数据中未出现的基准数据项的 field_name
    """
    missing_field_names = []
    for code, data in baseline_data.items():
        if code not in input_data:
            # Assumes 'field_name' key exists in the data dictionary
            missing_field_names.append(data['field_name'])
    return missing_field_names



def filter_missing_field_h2_names(baseline_data, input_code_array):
    """
    过滤出在基准数据中存在，但在输入数据中没有出现的字段名称。

    :param baseline_data: 字典，包含基准数据，格式为 {code: field_data}
    :param input_code_array: 字典，包含要检查的数据，格式应与 [code1,code2] 相同
    :return: 列表，包含那些在输入数据中未出现的基准数据项的 field_name
    """
    if input_code_array == "N/A":
        return ""
    missing_field_names = []
    for code, data in baseline_data.items():
        if code not in input_code_array:
            # Assumes 'field_name' key exists in the data dictionary
            missing_field_names.append(data['field_name'])
    return missing_field_names


def fill_placeholder_template(template_path, output_path, project, field_dict):
    """
    生成 Word 文档，替换正文、表格、页眉、页脚占位符，并处理图片替换。

    :param template_path: 原始 Word 模板路径
    :param output_path: 生成 Word 文档的目标路径
    :param project: 包含项目信息的对象
    :param field_list: 字段列表，包含 code 和 value
    """

    # **创建临时目录**
    temp_dir = output_path.replace(".docx", "_temp")
    DocxProcessor.unzip_docx(template_path, temp_dir)  # 解压原始 .docx
    # **项目信息映射**
    project_placeholders = {
        "{{project_model}}": project.project_model or "N/A",
        "{{project_name}}": project.project_name or "N/A",
        "{{project_type}}": project.project_type or 'N/A',
        "{{file_number}}": project.file_number or "N/A",
        "{{product_number}}": project.product_number or "N/A",
        "{{project_level}}": project.project_level or "N/A",
    }

    # **合并所有占位符**
    all_placeholders = {**project_placeholders, **field_dict}

    # **替换正文和表格**
    DocxProcessor.replace_docx_text(os.path.join(temp_dir, "word/document.xml"), all_placeholders)

    # **替换页眉和页脚**
    for file in os.listdir(os.path.join(temp_dir, "word")):
        if file.startswith("header") or file.startswith("footer"):
            DocxProcessor.replace_docx_text(os.path.join(temp_dir, f"word/{file}"), all_placeholders)

    print("✅ 文本占位符替换完成.")

    # **图片处理**

    # 获取图片路径
    dimensions_url = field_dict.get("{{dimensions}}")
    circuit_diagram_filename = field_dict.get("{{circuit_diagram}}")

    # 计算目标图片路径
    IMAGE_RM = os.path.join(app.config['IMAGES_FOLDER'], os.path.basename(dimensions_url)) if dimensions_url else None
    EMF_RM = os.path.join(app.config['EMF_FOLDER'], circuit_diagram_filename) if circuit_diagram_filename else None

    # **处理图片替换或删除**
    replacements_dict = {}

    if dimensions_url:
        replacements_dict["image1.png"] = IMAGE_RM

    if circuit_diagram_filename:
        replacements_dict["image2.emf"] = EMF_RM


    # **执行图片替换**
    if replacements_dict:
        DocxProcessor.replace_images_in_docx(os.path.join(temp_dir, "word", "media"), replacements_dict)
        print(f"✅ 已替换 {len(replacements_dict)} 张图片.")

    # **重新压缩 .docx**
    DocxProcessor.zip_docx(temp_dir, output_path)
    shutil.rmtree(temp_dir)  # 删除临时文件夹
    print(f"✅ 文档生成成功: {output_path}")



