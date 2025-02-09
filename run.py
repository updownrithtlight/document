from waitress import serve
from app import app
import os
import logging
from logging.handlers import RotatingFileHandler

if __name__ == "__main__":
    # 设置日志文件路径
    log_file = os.path.join(os.path.dirname(__file__), "server.log")
    error_file = os.path.join(os.path.dirname(__file__), "error.log")

    # 配置全局日志记录器
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # 普通日志处理器
    log_handler = RotatingFileHandler(log_file, maxBytes=1000000, backupCount=5, encoding="utf-8")
    log_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(log_handler)

    # 错误日志处理器
    error_handler = RotatingFileHandler(error_file, maxBytes=1000000, backupCount=5, encoding="utf-8")
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(error_handler)

    # 捕获 Waitress 日志
    waitress_logger = logging.getLogger("waitress")
    waitress_logger.setLevel(logging.INFO)
    waitress_logger.addHandler(log_handler)  # 添加普通日志处理器
    waitress_logger.addHandler(error_handler)  # 添加错误日志处理器

    # 服务配置
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", 5000))

    # 启动 Waitress
    serve(app, host=host, port=port, threads=8)
