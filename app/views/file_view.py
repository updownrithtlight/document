from flask import Blueprint
from app.controllers import file_controller

file_bp = Blueprint('file', __name__, url_prefix='/api/file')

@file_bp.route('/upload', methods=['POST'])
def upload():
    return file_controller.upload_file()

@file_bp.route('/download/<filename>', methods=['GET'])
def download(filename):
    return file_controller.download_file(filename)

@file_bp.route('/delete/<filename>', methods=['DELETE'])
def delete(filename):
    return file_controller.delete_file(filename)


@file_bp.route('/preview/<filename>', methods=['GET'])
def preview(filename):
    return file_controller.preview_file(filename)
