import os
import subprocess
import sys
from flask import Blueprint, request, send_file, jsonify

from app import app
from app.exceptions.exceptions import CustomAPIException
from app.controllers.word_controller import generate_document as tech_manual_document
from app.controllers.word_product_spec_controller import generate_document as product_spec_document
from app.controllers.excel_controller import generate_excel_local
from app.models.result import ResponseTemplate

office_file_bp = Blueprint('office_file', __name__, url_prefix='/api/office_file')

# 确保 UNO 模块可以找到
LIBREOFFICE_PROGRAM_PATH = app.config['LIBREOFFICE_LIB_PATH']

# LibreOffice 配置
LIBREOFFICE_PATH = app.config['LIBREOFFICE_PATH']
OUTPUT_FOLDER = app.config['OUTPUT_FOLDER']
def convert_to_pdf_with_command(file_path):
    """
    使用 LibreOffice 的命令行模式将文件转换为 PDF
    """
    abs_path = os.path.abspath(file_path)
    output_dir = os.path.dirname(file_path)
    output_pdf_path = os.path.splitext(abs_path)[0] + ".pdf"

    try:
        # 构造命令
        command = [
            LIBREOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            abs_path
        ]
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 转换失败: {result.stderr}")

        if not os.path.exists(output_pdf_path):
            raise FileNotFoundError(f"转换后的 PDF 文件未找到: {output_pdf_path}")

        app.logger.info(f"✅ PDF 文件已生成: {output_pdf_path}")
        return output_pdf_path
    except Exception as e:
        app.logger.error(f"❌ 文件转换为 PDF 失败: {e}")
        raise RuntimeError(f"文件转换为 PDF 失败: {e}")

# 📌 **生成技术手册**
@office_file_bp.route("/generate/tech-manual/<int:project_id>", methods=["GET"])
def generate_tech_manual_file(project_id):
    output_word_path,output_file_name = tech_manual_document(project_id)


    # **转换为 PDF**
    output_pdf_path = convert_to_pdf_with_command(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)


# 📌 **生成产品规格书**
@office_file_bp.route("/generate/product_spec/<int:project_id>", methods=["GET"])
def generate_product_spec_file(project_id):
    output_word_path,output_file_name = product_spec_document(project_id)

    # **转换为 PDF**
    output_pdf_path = convert_to_pdf_with_command(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)


@office_file_bp.route("/generate/excel-bom/<int:project_id>", methods=["GET"])
def generate_excel_file(project_id):
    output_excel_path = generate_excel_local(project_id)
    # **转换为 PDF**
    output_pdf_path = convert_to_pdf_with_command(output_excel_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)

@office_file_bp.route("/preview/<filename>", methods=["GET"])
def preview_file(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, mimetype="application/pdf")
    return ResponseTemplate.error("文件不存在", 404)



