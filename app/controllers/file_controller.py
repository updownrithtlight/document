import os
from flask import jsonify, request, send_file
from werkzeug.utils import secure_filename
from urllib.parse import quote
from datetime import datetime
from app.models.result import ResponseTemplate
from config import Config

# 获取文件存储路径
IMAGES_FOLDER = Config.IMAGES_FOLDER
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'}

# 确保上传目录存在
if not os.path.exists(IMAGES_FOLDER):
    os.makedirs(IMAGES_FOLDER)


def allowed_file(filename):
    """检查文件后缀是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def upload_file():
    """处理文件上传"""
    if "file" not in request.files:
        return ResponseTemplate.error("未找到文件")

    file = request.files["file"]

    if file.filename == "":
        return ResponseTemplate.error("未选择文件")

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)

        # 避免重名，可选：加时间戳
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        unique_filename = f"{timestamp}_{filename}"

        file_path = os.path.join(IMAGES_FOLDER, unique_filename)
        file.save(file_path)

        # 生成文件 URL
        file_url = f"http://localhost:5000/api/file/preview/{unique_filename}"

        return ResponseTemplate.success(
            message="文件上传成功",
            data={"fileName": unique_filename, "url": file_url}
        )

    return ResponseTemplate.error("文件上传失败，不支持的文件类型")


def download_file(filename):
    """处理文件下载"""
    file_path = os.path.join(IMAGES_FOLDER, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    encoded_file_name = quote(filename)
    response = send_file(file_path, as_attachment=True)
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_file_name}"
    return response


def delete_file(filename):
    """处理文件删除"""
    file_path = os.path.join(IMAGES_FOLDER, filename)

    if os.path.exists(file_path):
        os.remove(file_path)
        return ResponseTemplate.success(message="文件删除成功", data={"fileName": "","url":""})

    return ResponseTemplate.error("文件上传失败")


def preview_file(filename):
    """预览文件（适用于图片或文本文件）"""
    file_path = os.path.join(IMAGES_FOLDER, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    return send_file(file_path)
