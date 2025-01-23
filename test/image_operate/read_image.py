
import zipfile
import os

docx_path = 'imagetest.docx'

with zipfile.ZipFile(docx_path, 'r') as zip_ref:
    # 列出所有文件
    all_files = zip_ref.namelist()
    # 过滤出word/media文件夹中的文件
    media_files = [f for f in all_files if f.startswith('word/media/')]
    print(media_files)
