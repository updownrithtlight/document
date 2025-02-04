import json
from datetime import datetime
from flask import jsonify, send_file, request
from flask_jwt_extended import jwt_required
from urllib.parse import quote
from app import app
from app.controllers.project_field_controller import get_list_by_project_id
from app.controllers.field_definition_controller import get_fields_by_code
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
        # **提取参数**
        today = datetime.today()

        # 生成格式化的日期字符串
        formatted_date = today.strftime("%y%m%d")

        print(formatted_date)
        # **生成文件路径**
        output_file_name = f"{project.project_model}技术说明书 {formatted_date}.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file_name)
        # **转换 field_list 为字典**
        data_map = {item['code']: item for item in project_field_list if item.get('code') is not None}
        print(f"📌 解析字段完成，共 {len(data_map)} 个字段.")
        # **填充 Word 模板**
        fill_placeholder_template(TECHNICAL_TEMPLATE_PATH, output_path, project, data_map)

        try:
            doc = Document(output_path)
        except Exception as e:
            print("❌ 加载文档失败：", e)
            return None

        data_source_map = get_fields_by_code()

        target_titles = filter_missing_field_names(data_source_map, data_map)
        print(f"📌 需要删除的标题: {target_titles}")

        # **删除未出现的标题**
        for title in target_titles:
            WordTocTool.delete_section_by_title(doc, title)

        # **保存删除后的文档**
        doc.save(output_path)  # ✅ 这里确保删除的内容被保存

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


def fill_placeholder_template(template_path, output_path, project, data_map):
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
    print("🔍 解压完成，开始处理字段替换...")

    # **安全获取 manufacturing_process 并解析**
    manufacturing_process_data = data_map.get("manufacturing_process", {}).get("custom_value", "N/A")
    try:
        data_list = json.loads(manufacturing_process_data) if manufacturing_process_data not in ["N/A", None,
                                                                                                 ""] else []
    except json.JSONDecodeError:
        data_list = []

    formatted_str = "、".join(data_list) if data_list else "N/A"
    print(formatted_str)

    # **构建字段映射**
    field_dict = {
        "{{operating_temp}}": data_map.get("operating_temp", {}).get("custom_value", "N/A"),
        "{{storage_temp}}": data_map.get("storage_temp", {}).get("custom_value", "N/A"),
        "{{housing_material}}": data_map.get("housing_material", {}).get("custom_value", "N/A"),
        "{{manufacturing_process}}": formatted_str,
        "{{weight}}": data_map.get("weight", {}).get("custom_value", "N/A"),
        "{{input_terminal}}": data_map.get("input_terminal", {}).get("custom_value", "N/A"),
        "{{output_terminal}}": data_map.get("output_terminal", {}).get("custom_value", "N/A"),
    }

    print(field_dict)

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
    def safe_get(data, key, sub_key):
        """安全获取嵌套字典值，避免 KeyError"""
        return data.get(key, {}).get(sub_key)

    # 获取图片路径
    dimensions_url = safe_get(data_map, "dimensions", "custom_value")
    circuit_diagram_filename = safe_get(data_map, "circuit_diagram", "custom_value")

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
