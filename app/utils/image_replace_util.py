import zipfile
import os
import shutil


def replace_images_in_docx(docx_path, replacements, output_path=None):
    """
    用新的图片替换 docx 文件中的图片

    :param docx_path:      原始 docx 文件路径
    :param replacements:   dict，键为需要被替换的图片文件名（例如 'image1.png'），值为新的图片路径
    :param output_path:    替换后的 docx 文件输出路径，为 None 时自动在同目录生成一个新文件
    :return:               替换后 docx 文件的路径
    """
    # 如果未指定输出路径，则在原文件同目录下生成一个新文件
    if output_path is None:
        base, ext = os.path.splitext(docx_path)
        output_path = f"{base}_replaced{ext}"

    # 1. 创建一个临时文件夹用于解压
    tmp_dir = "temp_docx_extract"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # 2. 解压 docx 到临时文件夹
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)

    # 3. 替换图片（仅限同名/同扩展名）
    media_dir = os.path.join(tmp_dir, "word", "media")
    if not os.path.exists(media_dir):
        # 如果没有 word/media 文件夹，说明里面没有图片
        print("文档中未找到图片文件夹 word/media")
    else:
        for old_image_name, new_image_path in replacements.items():
            old_image_full_path = os.path.join(media_dir, old_image_name)
            if os.path.exists(old_image_full_path):
                # 删除旧图，然后复制新图过去
                os.remove(old_image_full_path)
                shutil.copy(new_image_path, old_image_full_path)
                print(f"已替换: {old_image_name} -> {new_image_path}")
            else:
                print(f"未找到对应图片: {old_image_name}，跳过替换。")

    # 4. 重新压缩成 docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for root, dirs, files in os.walk(tmp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # 让文件在 zip 中的相对路径与解压后对应
                arcname = os.path.relpath(file_path, tmp_dir)
                zip_out.write(file_path, arcname=arcname)

    # 5. 清理临时文件夹
    shutil.rmtree(tmp_dir)

    return output_path


if __name__ == "__main__":
    # 示例：假设我们要替换 Word 文档里的两张图片 image1.png 和 image2.png
    docx_file = r"C:\noneSystem\bj\document\uploads\template\technical_document_template.docx"
    # 字典形式：{"原图片文件名": "新图片的实际路径"}

    # replacements_dict = {
    #     "image1.png": r"C:\new_images\new_image.png",  # 替换 PNG
    #     "image2.emf": r"C:\new_images\new_image.emf",  # 替换 EMF
    # }

    replacements_dict = {
        "image1.png": "new.png",
    }
    output_file_name="temxxxxxplate.docx"
    output_path = rf"C:\noneSystem\bj\document\uploads\output\{output_file_name}"
    new_docx = replace_images_in_docx(docx_file, replacements_dict,output_path)
    print("新的文档输出路径：", new_docx)
