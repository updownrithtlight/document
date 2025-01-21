from flask import jsonify
from werkzeug.exceptions import HTTPException
from app.models.result import ResponseTemplate


class CustomAPIException(Exception):
    """自定义 API 异常"""
    def __init__(self, message="Internal Server Error", status_code=500, data=None):
        super().__init__()
        self.message = message
        self.status_code = status_code
        self.data = data

    def to_response(self):
        """转换为标准 ResponseTemplate 错误响应"""
        return ResponseTemplate.error(message=self.message, status_code=self.status_code, data=self.data)


def register_error_handlers(app):
    """注册全局异常处理器"""

    @app.errorhandler(CustomAPIException)
    def handle_custom_exception(error):
        return error.to_response()

    @app.errorhandler(HTTPException)
    def handle_http_exception(error):
        return ResponseTemplate.error(message=error.description, status_code=error.code)

    @app.errorhandler(Exception)
    def handle_generic_exception(error):
        """捕获所有未处理的异常"""
        return ResponseTemplate.error(message="Internal Server Error", status_code=500)
