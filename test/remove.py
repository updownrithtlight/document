import os
import subprocess
import time
from docx import Document
import sys

# 将 LibreOfficePortable 的 program 目录加入 sys.path，以便导入 UNO 模块
libreoffice_program_path = r"C:\noneSystem\bj\document\LibreOfficePortable\App\libreoffice\program"
if libreoffice_program_path not in sys.path:
    sys.path.append(libreoffice_program_path)
    print(f"已将 '{libreoffice_program_path}' 添加到 sys.path")

try:
    import uno
    from com.sun.star.beans import PropertyValue
    print("UNO 模块导入成功！")
except ImportError as e:
    print("导入 UNO 模块失败，请确保使用 LibreOffice 自带的 Python 或正确配置 UNO 环境。", e)
    # 如果导入失败，可以选择退出或者继续（但 UNO 功能会失效）
    # exit(1)

def start_libreoffice_portable():
    """
    自动启动 LibreOfficePortable 的 headless UNO 服务。
    注意：请确保 LibreOfficePortable 的路径正确。
    """
    LIBREOFFICE_PATH = r"C:\noneSystem\bj\document\LibreOfficePortable\App\libreoffice\program\soffice.exe"
    cmd = [
        LIBREOFFICE_PATH,
        "--headless",
        '--accept=socket,host=localhost,port=2002;urp;',
        "--norestore",
        "--nolockcheck",
        "--nodefault"
    ]
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        time.sleep(3)  # 等待 UNO 服务启动
        print("LibreOfficePortable headless 服务已启动。")
        return proc
    except Exception as e:
        print("启动 LibreOfficePortable 失败：", e)
        return None

def delete_section_by_title(doc, target_title, heading_level="Heading 1"):
    """
    删除 Word 文档中特定标题及其后续内容，直到遇到下一个相同级别的标题。
    """
    def delete_paragraph(paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        paragraph._p = paragraph._element = None

    delete_flag = False
    for para in doc.paragraphs[:]:
        if para.style.name == heading_level:
            if para.text.strip() == target_title:
                delete_flag = True
            else:
                delete_flag = False
        if delete_flag:
            delete_paragraph(para)

def update_toc_with_uno(doc_path, host="localhost", port=2002):
    """
    利用 LibreOffice UNO 接口直接加载 docx 文件并更新目录（TOC）。
    """
    abs_path = os.path.abspath(doc_path)
    file_url = uno.systemPathToFileUrl(abs_path)
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
    try:
        ctx = resolver.resolve(f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext")
    except Exception as e:
        print("连接到 LibreOfficePortable UNO 服务失败，请确认服务已启动：", e)
        return

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    load_props = tuple()
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)
    if not doc:
        print("加载文档失败。")
        return

    try:
        doc.updateAll()
    except Exception as e:
        print("更新字段时出错：", e)
    store_props = (PropertyValue("FilterName", 0, "MS Word 2007 XML", 0),)
    try:
        doc.storeToURL(file_url, store_props)
    except Exception as e:
        print("保存文档时出错：", e)
    finally:
        doc.close(True)
    print("TOC 更新成功。")

def main():
    libreoffice_proc = start_libreoffice_portable()
    if not libreoffice_proc:
        print("无法启动 LibreOfficePortable，程序终止。")
        return

    try:
        # 模板文件路径（不修改此文件）
        template_path = r"C:\noneSystem\bj\document\test\technical_document_template.docx"
        if not os.path.exists(template_path):
            print("模板文件不存在：", template_path)
            return

        # 构造新的输出文件路径
        new_doc_path = template_path.replace(".docx", "_modified.docx")
        print("加载模板文档进行章节删除...")
        try:
            doc = Document(template_path)
        except Exception as e:
            print("加载文档失败：", e)
            return

        print("删除标题 '插入损耗特性' 及其后续内容...")
        delete_section_by_title(doc, "插入损耗特性")
        try:
            doc.save(new_doc_path)
            print(f"文档修改成功，保存至：{new_doc_path}")
        except Exception as e:
            print("保存文档失败：", e)
            return

        print("调用 UNO 接口更新 TOC ...")
        update_toc_with_uno(new_doc_path)
    finally:
        libreoffice_proc.terminate()
        print("LibreOfficePortable headless 服务已关闭。")

if __name__ == "__main__":
    main()
