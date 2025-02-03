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

if LIBREOFFICE_PROGRAM_PATH not in sys.path:
    sys.path.append(LIBREOFFICE_PROGRAM_PATH)
    app.logger.info(f"✅ 已将 LibreOffice `program` 目录添加到 sys.path: {LIBREOFFICE_PROGRAM_PATH}")

try:
    import uno
    from com.sun.star.beans import PropertyValue
    app.logger.info("✅ UNO 模块导入成功！")
except ImportError as e:
    app.logger.error("❌ UNO 模块导入失败，请检查 LibreOffice 是否正确安装: %s", e)
    sys.exit(1)

# LibreOffice 配置
LIBREOFFICE_PATH = app.config['LIBREOFFICE_PATH']
OUTPUT_FOLDER = app.config['OUTPUT_FOLDER']


# 📌 **更新 Word 文档中的目录 (TOC)**
def update_toc_with_uno(doc_path):
    """
    使用 UNO 接口更新 Word 文档的目录（TOC）。
    """
    abs_path = os.path.abspath(doc_path)
    file_url = uno.systemPathToFileUrl(abs_path)

    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)

    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        app.logger.error("❌ 连接到 LibreOffice UNO 失败，请确认 LibreOffice 服务已启动: %s", e)
        return {"error": "LibreOffice UNO 连接失败", "details": str(e)}

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 加载 Word 文档
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    if not doc:
        app.logger.error("❌ 加载 Word 文档失败: %s", doc_path)
        return {"error": "加载 Word 文档失败"}

    try:
        # 更新所有目录字段
        doc.updateAll()
    except Exception as e:
        app.logger.error("❌ 更新 Word 目录 (TOC) 失败: %s", e)
        return {"error": "更新 TOC 失败", "details": str(e)}

    # 保存文档
    store_props = (PropertyValue("FilterName", 0, "MS Word 2007 XML", 0),)
    try:
        doc.storeToURL(file_url, store_props)
    except Exception as e:
        app.logger.error("❌ 保存 Word 文档失败: %s", e)
        return {"error": "保存 Word 失败", "details": str(e)}
    finally:
        doc.close(True)

    app.logger.info("✅ Word 目录 (TOC) 已更新: %s", doc_path)
    return {"success": True}


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
