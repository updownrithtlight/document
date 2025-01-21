import os
from datetime import timedelta

from dotenv import load_dotenv
load_dotenv()

class Config:
    TIMEZONE = 'Asia/Shanghai'
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'your_secret_key_here'
    # SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///database.db'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    BCRYPT_LOG_ROUNDS = 12
    JWT_SECRET_KEY = os.environ.get('JWT_SECRET_KEY') or 'your_jwt_secret_key_here'
    JWT_ALGORITHM = 'HS256'
    CORS_HEADERS = 'Content-Type'
    JWT_HEADER_NAME = 'Authorization'
    JWT_HEADER_TYPE = 'Bearer'
    JWT_ACCESS_TOKEN_EXPIRES = timedelta(days=1)  # Access Token 的过期时间
    JWT_REFRESH_TOKEN_EXPIRES = timedelta(days=30)  # Refresh Token 的过期时间
    JWT_TOKEN_LOCATION = ['headers', 'cookies']  # 定义 JWT 的读取位置
    JWT_COOKIE_CSRF_PROTECT = False  # 如果你用 cookie 存储 token，是否启用 CSRF 保护
    JWT_COOKIE_SECURE = False  # 是否强制在 HTTPS 下使用
    JWT_COOKIE_SAMESITE = 'Lax'  # Cookie 的 SameSite 策略
    # 文件上传目录配置
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 项目根目录
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')  # 上传文件的目录路径
    TEMPLATE_FOLDER = os.path.join(UPLOAD_FOLDER, 'template')  # 存放模板文件的子目录
    OUTPUT_FOLDER = os.path.join(UPLOAD_FOLDER, 'output')  # 存放生成文件的子目录
    FIELD_IMAGES_FOLDER = os.path.join(UPLOAD_FOLDER, 'field_images')  # 存放模板文件的子目录
    IMAGES_FOLDER = os.path.join(UPLOAD_FOLDER, 'images')  # 存放模板文件的子目录



