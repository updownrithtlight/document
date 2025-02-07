import os
import zipfile
import shutil
from lxml import etree

class DocxProcessor:
    @staticmethod
    def unzip_docx(docx_path, extract_to):
        """解压 Word 文档"""
        if not os.path.exists(extract_to):
            os.makedirs(extract_to)
        with zipfile.ZipFile(docx_path, 'r') as docx:
            docx.extractall(extract_to)

    @staticmethod
    def zip_docx(folder_path, output_path):
        """重新打包 .docx 文件"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, folder_path)
                    docx.write(file_path, arcname)

    @staticmethod
    def replace_docx_text(xml_path, replacements):
        """使用 lxml 替换 XML 文本"""
        with open(xml_path, "r", encoding="utf-8") as f:
            xml_content = f.read()

        # 替换占位符
        for placeholder, value in replacements.items():
            xml_content = xml_content.replace(placeholder, value)

        with open(xml_path, "w", encoding="utf-8") as f:
            f.write(xml_content)

    @staticmethod
    def replace_images_in_docx(media_folder_path, replacements):
        """
        用新的图片替换 docx 文件中的图片
        :param media_folder_path: 原始 docx 解压后 media 文件夹路径
        :param replacements: dict，键为需要被替换的图片文件名（例如 'image1.png'），值为新的图片路径
        """
        if not os.path.exists(media_folder_path):
            print("文档中未找到图片文件夹 word/media")
            return

        for old_image_name, new_image_path in replacements.items():
            old_image_full_path = os.path.join(media_folder_path, old_image_name)
            if os.path.exists(old_image_full_path):
                os.remove(old_image_full_path)
                shutil.copy(new_image_path, old_image_full_path)
                print(f"已替换: {old_image_name} -> {new_image_path}")
            else:
                print(f"未找到对应图片: {old_image_name}，跳过替换。")
