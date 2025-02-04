import os
import win32com.client as win32
from docx import Document


def delete_section_by_title(doc, target_title, heading_level="Heading 1"):
    """
    删除 Word 文档中，从指定标题开始直到下一个同级标题出现之前的所有段落。
    注意：此方法依赖 python‑docx 内部对象，未来版本可能需要调整。
    """

    def delete_paragraph(paragraph):
        element = paragraph._element
        element.getparent().remove(element)
        # 将内部对象置空，确保段落被删除
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


def adjust_toc_tab_stops(toc, new_tab_position=400):
    """
    针对 TOC 中的段落进行手动调整：
      - 对于样式为 "TOC 1" 的段落，清除原有制表符后添加一个新的制表符，
        位置设为 new_tab_position（单位：点），右对齐，leader 为点线；
      - 同时设置这些段落的字体为“宋体”，字号设置为 12pt（小四）。
    如果某些段落不支持访问 ParagraphFormat，则直接跳过。
    """
    constants = win32.constants
    toc_range = toc.Range
    for para in toc_range.Paragraphs:
        try:
            # 判断是否存在 ParagraphFormat 属性
            if not hasattr(para, "ParagraphFormat"):
                continue
            # 获取段落的样式名称（部分 Word 版本使用 NameLocal）
            style = para.Style
            style_name = ""
            try:
                style_name = style.NameLocal.strip()
            except Exception:
                try:
                    style_name = style.Name.strip()
                except Exception:
                    pass
            # 针对 "TOC 1" 级别的段落进行调整
            if style_name == "TOC 1":
                try:
                    para.ParagraphFormat.TabStops.ClearAll()
                    para.ParagraphFormat.TabStops.Add(Position=new_tab_position,
                                                      Alignment=constants.wdAlignTabRight,
                                                      Leader=constants.wdTabLeaderDots)
                    # 设置字体为宋体，小四（12pt）
                    para.Range.Font.Name = "宋体"
                    para.Range.Font.Size = 12
                except Exception:
                    continue
        except Exception:
            continue


def set_toc_heading_style(word_app):
    """
    将目录标题（TOC Heading 样式）的字体设置为宋体、小四（12pt）。
    如果文档中没有“TOC Heading”样式，该函数会捕获异常并提示。
    """
    doc = word_app.ActiveDocument
    try:
        toc_heading = doc.Styles("TOC Heading")
        toc_heading.Font.Name = "宋体"
        toc_heading.Font.Size = 12  # 小四约 12pt
        toc_heading.ParagraphFormat.SpaceBefore = 0
        toc_heading.ParagraphFormat.SpaceAfter = 0
        print("TOC Heading 样式已更新为：宋体，小四号")
    except Exception as e:
        print("无法更新 TOC Heading 样式：", e)


def update_toc_via_word(doc_path, template_path=None):
    """
    利用 Microsoft Word COM 自动化接口更新文档目录（TOC），
    并在更新后针对 TOC 1 级别的条目手动设置制表符及字体，
    同时更新目录标题（TOC Heading 样式）的格式为宋体、小四号。
    """
    abs_path = os.path.abspath(doc_path)
    print("启动 Word 应用程序...")
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 后台运行

    try:
        print("打开文档：", abs_path)
        doc = word.Documents.Open(abs_path)
    except Exception as e:
        print("❌ 打开文档失败：", e)
        word.Quit()
        return False

    try:
        if doc.TablesOfContents.Count > 0:
            toc = doc.TablesOfContents(1)
            print("更新目录 TOC ...")
            toc.Update()
            # 更新后手动调整 TOC 1 级别的制表符及字体设置
            adjust_toc_tab_stops(toc, new_tab_position=400)
        else:
            print("文档中没有目录 TOC")

        # 设置目录标题（TOC Heading）的格式
        # set_toc_heading_style(word)

        print("保存文档...")
        doc.Save()
    except Exception as e:
        print("❌ 更新目录失败：", e)
        doc.Close(False)
        word.Quit()
        return False

    try:
        doc.Close(False)
        print("关闭文档成功。")
    except Exception as e:
        print("关闭文档时出现异常：", e)
    finally:
        word.Quit()
        print("Word 应用程序已退出。")
    return True


def main():
    # 模板文档路径，请根据实际情况修改
    template_path = r"C:\noneSystem\bj\document\test\technical_document_template.docx"
    if not os.path.exists(template_path):
        print("❌ 模板文件不存在：", template_path)
        return

    new_doc_path = template_path.replace(".docx", "_modified.docx")

    print("📌 正在加载模板文档并删除指定章节...")
    try:
        doc = Document(template_path)
    except Exception as e:
        print("❌ 加载文档失败：", e)
        return

    print("📌 正在删除标题 '插入损耗特性' 及其后续内容...")
    delete_section_by_title(doc, "插入损耗特性")
    delete_section_by_title(doc, "产品功能")

    try:
        doc.save(new_doc_path)
        print(f"✅ 文档修改成功，已保存到：{new_doc_path}")
    except Exception as e:
        print("❌ 保存文档失败：", e)
        return

    print("📌 正在调用 Word COM 自动化接口更新 TOC ...")
    if update_toc_via_word(new_doc_path):
        print("✅ TOC 更新成功")
    else:
        print("❌ TOC 更新失败")


if __name__ == "__main__":
    main()
