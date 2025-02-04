import os
import sys
from flask import Blueprint, request, send_file, jsonify
from app import app
from app.exceptions.exceptions import CustomAPIException
from app.controllers.office_document_controller import generate_tech_manual, generate_product_spec
from app.controllers.excel_controller import generate_excel_local
from app.models.result import ResponseTemplate

office_file_bp = Blueprint('office_file', __name__, url_prefix='/api/office_file')

# 确保 UNO 模块可以找到
LIBREOFFICE_PROGRAM_PATH = app.config['LIBREOFFICE_LIB_PATH']


# LibreOffice 配置
LIBREOFFICE_PATH = app.config['LIBREOFFICE_PATH']
OUTPUT_FOLDER = app.config['OUTPUT_FOLDER']

# 🔄 **使用 UNO 进行 PDF 转换**
def convert_to_pdf(file_path):
    """
    使用 UNO 将 Word/Excel 转换为 PDF
    """
    abs_path = os.path.abspath(file_path)
    file_url = uno.systemPathToFileUrl(abs_path)
    output_pdf_path = os.path.join(OUTPUT_FOLDER, os.path.splitext(os.path.basename(file_path))[0] + ".pdf")
    output_pdf_url = uno.systemPathToFileUrl(output_pdf_path)

    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)

    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        app.logger.error("❌ 连接到 LibreOffice UNO 失败: %s", e)
        return None

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 加载文档
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    if not doc:
        app.logger.error("❌ 加载文档失败: %s", file_path)
        return None

    try:
        # 另存为 PDF
        store_props = (PropertyValue("FilterName", 0, "writer_pdf_Export", 0),)
        doc.storeToURL(output_pdf_url, store_props)
    except Exception as e:
        app.logger.error("❌ 转换为 PDF 失败: %s", e)
        return None
    finally:
        doc.close(True)

    app.logger.info("✅ 文件已转换为 PDF: %s", output_pdf_path)
    return output_pdf_path


# 📌 **生成技术手册**
@office_file_bp.route("/generate/tech-manual/<int:project_id>", methods=["GET"])
def generate_tech_manual_file(project_id):
    output_word_path = generate_tech_manual(project_id)

    # **更新 Word 目录**
    update_result = update_toc_with_uno(output_word_path)
    if "error" in update_result:
        raise CustomAPIException(update_result["error"], 500)

    # **转换为 PDF**
    output_pdf_path = convert_to_pdf(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)


# 📌 **生成产品规格书**
@office_file_bp.route("/generate/product_spec/<int:project_id>", methods=["GET"])
def generate_product_spec_file(project_id):
    output_word_path = generate_product_spec(project_id)

    # **更新 Word 目录**
    update_result = update_toc_with_uno(output_word_path)
    if "error" in update_result:
        raise CustomAPIException(update_result["error"], 500)

    # **转换为 PDF**
    output_pdf_path = convert_to_pdf(output_word_path)
    if not output_pdf_path:
        raise CustomAPIException("文件转换失败", 500)

    return os.path.basename(output_pdf_path)
