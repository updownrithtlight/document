import os
import subprocess
from flask import Blueprint, request, send_file, jsonify
from app import app
from app.exceptions.exceptions import CustomAPIException
from app.controllers.office_document_controller import generate_tech_manual,generate_product_spec
from app.controllers.excel_controller import generate_excel_local
from app.models.result import ResponseTemplate

office_file_bp = Blueprint('office_file', __name__, url_prefix='/api/office_file')

# LibreOffice Portable 路径
LIBREOFFICE_PATH = app.config['LIBREOFFICE_PATH']

OUTPUT_FOLDER = app.config['OUTPUT_FOLDER']


# 🔄 转换文件为 PDF
def convert_to_pdf(file_path):
    if not os.path.exists(file_path):
        return None

    cmd = [
        LIBREOFFICE_PATH,
        "--headless", "--convert-to", "pdf",
        "--outdir", OUTPUT_FOLDER, file_path
    ]

    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        pdf_filename = os.path.splitext(os.path.basename(file_path))[0] + ".pdf"
        return os.path.join(OUTPUT_FOLDER, pdf_filename)
    except subprocess.CalledProcessError as e:
        print("文件转换错误:", e)
        return None


# 📌 提供文件预览
@office_file_bp.route("/generate/tech-manual/<int:project_id>", methods=["GET"])
def generate_tech_manual_file(project_id):
    output_word_path = generate_tech_manual(project_id)
    # **转换为 PDF**
    output_pdf_path = convert_to_pdf(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)


@office_file_bp.route("/generate/product_spec/<int:project_id>", methods=["GET"])
def generate_product_spec_file(project_id):
    output_word_path = generate_product_spec(project_id)
    # **转换为 PDF**
    output_pdf_path = convert_to_pdf(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)



@office_file_bp.route("/preview/<filename>", methods=["GET"])
def preview_file(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, mimetype="application/pdf")
    return ResponseTemplate.error("文件不存在", 404)




@office_file_bp.route("/generate/excel-bom/<int:project_id>", methods=["GET"])
def generate_excel_file(project_id):
    output_excel_path = generate_excel_local(project_id)
    # **转换为 PDF**
    output_pdf_path = convert_to_pdf(output_excel_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)