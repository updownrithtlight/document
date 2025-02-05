import json
from datetime import datetime
from flask import jsonify, send_file, request
from flask_jwt_extended import jwt_required
from urllib.parse import quote
from app import app
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
from app.utils.word_toc_tool import WordTocTool

# **模板文件路径**
TECHNICAL_TEMPLATE_PATH = os.path.join(app.config['TEMPLATE_FOLDER'], "technical_document_template.docx")
PRODUCT_SPECIFICATION_TEMPLATE_PATH = os.path.join(app.config['TEMPLATE_FOLDER'], "product_specification.docx")

@jwt_required()
def generate_tech_manual(project_id):
    """
    生成产品规范 Word 文档
    """
    try:
        project = Project.query.get(project_id)
        if not project:
            return jsonify({"error": "项目不存在"}), 404

        project_field_list = get_list_by_project_id(project_id)

        valid_parent_ids = [3, 4, 5, 6, 7, 8]

        # 1. 过滤出 parent_id 在 [3,4,5,6,7,8] 的列表
        filtered_part_1 = [
            item for item in project_field_list
            if item.get("parent_id") in valid_parent_ids
        ]
        print("过滤完",filtered_part_1)
        # 2. 过滤出 parent_id 不在 [3,4,5,6,7,8] 的列表
        filtered_part_2 = [
            item for item in project_field_list
            if item.get("parent_id") not in valid_parent_ids
        ]
        # **提取参数**
        today = datetime.today()

        # 生成格式化的日期字符串
        formatted_date = today.strftime("%y%m%d")

        print(formatted_date)
        # **生成文件路径**
        output_file_name = f"{project.project_model}技术说明书 {formatted_date}.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file_name)
        # **转换 field_list 为字典**

        cleaned_dict = build_cleaned_dict(filtered_part_2)

        print(cleaned_dict)
        # **填充 Word 模板**
        fill_placeholder_template(TECHNICAL_TEMPLATE_PATH, output_path, project, cleaned_dict)

        data_map = {item['code']: item for item in filtered_part_2 if item.get('code') is not None}

        try:
            doc = Document(output_path)
        except Exception as e:
            print("❌ 加载文档失败：", e)
            return None

        data_source_map = get_fields_by_code()

        target_titles = filter_missing_field_names(data_source_map, data_map)
        print(f"📌 需要删除的标题: {target_titles}")

        # **删除未出现的1级标题**
        for title in target_titles:
            WordTocTool.delete_section_by_title(doc, title)

        data_source_h2_map = get_fields_h2_by_code()
        environmental_characteristics = data_map.get("environmental_characteristics", {}).get("custom_value", "N/A")

        target_h2_titles = filter_missing_field_h2_names(data_source_h2_map, environmental_characteristics)
        if target_h2_titles != "":
            # **删除未出现的2级标题**
            for title in target_h2_titles:
                WordTocTool.delete_section_by_title2_or_higher(doc, title)
            # **保存删除后的文档**
        doc.save(output_path)  # ✅ 这里确保删除的内容被保存
        rows_to_add =  get_fields_by_project_id_parent_id(project_id, 44)

        new_row = [project.project_name, project.project_model,"1套", "粘贴标签、序列号、合格证"]
        rows_to_add.insert(0, new_row)
        # 3. 循环添加行
        for row_data in rows_to_add:
            add_row_with_auto_serial(doc, table_index=3, cell_values=row_data)

        # **保存删除后的文档**
        doc.save(output_path)  # ✅ 这里确保删除的内容被保存

        features = get_features(project_id=project_id)
        important_notes = get_important_notes(project_id=project_id)
        context = {}
        context.update(features)  # context 现在包含 {"features": [...]}
        context.update(important_notes)

        WordTocTool.fill_doc_with_features(output_path, context)

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
        raise CustomAPIException(e, 404)


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
    unzip_docx(template_path, temp_dir)  # 解压原始 .docx
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
    replace_docx_text(os.path.join(temp_dir, "word/document.xml"), all_placeholders)

    # **替换页眉和页脚**
    for file in os.listdir(os.path.join(temp_dir, "word")):
        if file.startswith("header") or file.startswith("footer"):
            replace_docx_text(os.path.join(temp_dir, f"word/{file}"), all_placeholders)

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
        replace_images_in_docx(os.path.join(temp_dir, "word", "media"), replacements_dict)
        print(f"✅ 已替换 {len(replacements_dict)} 张图片.")

    # **重新压缩 .docx**
    zip_docx(temp_dir, output_path)
    shutil.rmtree(temp_dir)  # 删除临时文件夹
    print(f"✅ 文档生成成功: {output_path}")


def unzip_docx(docx_path, extract_to):
    """ 解压 Word 文档 """
    with zipfile.ZipFile(docx_path, 'r') as docx:
        docx.extractall(extract_to)


def zip_docx(folder_path, output_path):
    """ 重新打包 .docx 文件 """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                docx.write(file_path, arcname)


def replace_docx_text(xml_path, replacements):
    """ 使用 lxml 替换 XML 文本 """
    with open(xml_path, "r", encoding="utf-8") as f:
        xml_content = f.read()

    # **替换占位符**
    for placeholder, value in replacements.items():
        xml_content = xml_content.replace(placeholder, value)

    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_content)



def replace_images_in_docx(docx_path, replacements):
    """
    用新的图片替换 docx 文件中的图片
    :param docx_path:      原始 docx 文件路径
    :param replacements:   dict，键为需要被替换的图片文件名（例如 'image1.png'），值为新的图片路径
    :param output_path:    替换后的 docx 文件输出路径，为 None 时自动在同目录生成一个新文件
    :return:               替换后 docx 文件的路径
    """

    # 3. 替换图片（仅限同名/同扩展名）
    if not os.path.exists(docx_path):
        # 如果没有 word/media 文件夹，说明里面没有图片
        print("文档中未找到图片文件夹 word/media")
    else:
        for old_image_name, new_image_path in replacements.items():
            old_image_full_path = os.path.join(docx_path, old_image_name)
            if os.path.exists(old_image_full_path):
                # 删除旧图，然后复制新图过去
                os.remove(old_image_full_path)
                shutil.copy(new_image_path, old_image_full_path)
                print(f"已替换: {old_image_name} -> {new_image_path}")
            else:
                print(f"未找到对应图片: {old_image_name}，跳过替换。")




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
        # **提取参数**


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
