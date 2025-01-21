import os
from flask import jsonify, send_file, request
from flask_jwt_extended import jwt_required
from urllib.parse import quote
from app import app
from app.controllers.project_field_controller import get_list_by_project_id
from app.exceptions.exceptions import CustomAPIException
from app.models.models import Project
import os
import zipfile
import shutil
from lxml import etree
from docx import Document

# **模板文件路径**
TEMPLATE_PATH = os.path.join(app.config['TEMPLATE_FOLDER'], "technical_document_template.docx")
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

        # **生成文件路径**
        output_file_name = f"{project.project_model}_技术说明书.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file_name)

        # **填充 Word 模板**
        fill_product_spec_template(TEMPLATE_PATH, output_path, project, project_field_list)

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
        raise CustomAPIException("Material not found in the project", 404)




def fill_product_spec_template(template_path, output_path, project, field_list):
    """
    生成 Word 文档，替换正文、表格、页眉、页脚占位符
    """
    temp_dir = output_path.replace(".docx", "_temp")  # 创建临时目录
    unzip_docx(template_path, temp_dir)  # 解压原始 .docx
    print("🔍 field_list 类型:", type(field_list))

    if field_list:
        print("🔍 field_list 第一个元素类型:", type(field_list[0]))
        print("🔍 field_list 第一个元素:", field_list[0])
    else:
        print("⚠️ field_list 为空列表")
    # **构建字段映射**
    field_dict = {
        f"{{{{POWER_{(field.get('code') or 'UNKNOWN').upper()}}}}}": field.get('value') or field.get('custom_value') or 'test'
        for field in field_list
    }

    project_placeholders = {
        "{{project_model}}": project.project_model,
        "{{project_name}}": project.project_name,
        "{{project_type}}": project.project_type or 'N/A',
        "{{working_temperature}}": project.working_temperature or 'N/A',
        "{{storage_temperature}}": project.storage_temperature or 'N/A',
        "{{file_number}}": project.file_number,
        "{{product_number}}": project.product_number,
        "{{project_level}}": project.project_level,
    }
    all_placeholders = {**project_placeholders, **field_dict}

    # **替换正文和表格**
    replace_docx_text(os.path.join(temp_dir, "word/document.xml"), all_placeholders)

    # **替换页眉和页脚**
    for file in os.listdir(os.path.join(temp_dir, "word")):
        if file.startswith("header") or file.startswith("footer"):
            replace_docx_text(os.path.join(temp_dir, f"word/{file}"), all_placeholders)

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
        # **提取参数**

        # **生成文件路径**
        output_file_name = f"{project.project_model}_产品规范.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file_name)

        # **填充 Word 模板**
        fill_product_spec_template(PRODUCT_SPECIFICATION_TEMPLATE_PATH, output_path, project, project_field_list)

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
