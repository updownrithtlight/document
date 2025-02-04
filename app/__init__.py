import sys
import subprocess
import time
import atexit
import logging
import pytz
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_jwt_extended import JWTManager, jwt_required
from flask_cors import CORS
from config import Config

app = Flask(__name__)
app.config.from_object(Config)
db = SQLAlchemy(app)

# 配置日志输出
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
logger.addHandler(logging.StreamHandler())  # 将日志输出到标准输出

bcrypt = Bcrypt(app)
jwt = JWTManager(app)
CORS(app, supports_credentials=True, resources={r"/*": {"origins": "http://localhost:3000"}}, expose_headers=["Content-Disposition"])

china_tz = pytz.timezone(app.config["TIMEZONE"])

def get_local_time(utc_time):
    """ 将 UTC 时间转换为中国标准时间（CST）"""
    return utc_time.replace(tzinfo=pytz.utc).astimezone(china_tz).strftime("%Y-%m-%d %H:%M:%S")

from app.routes import setup_routes
setup_routes(app)

# 自动创建表（仅开发环境）
with app.app_context():
    db.create_all()
    logger.info("✅ 数据库表已创建")

# 捕获 SQLAlchemy 的日志
logging.getLogger('sqlalchemy.engine').setLevel(logging.INFO)

# ============ 确保 UNO 模块可以找到 =============
LIBREOFFICE_PROGRAM_PATH = app.config['LIBREOFFICE_LIB_PATH']

if LIBREOFFICE_PROGRAM_PATH not in sys.path:
    sys.path.append(LIBREOFFICE_PROGRAM_PATH)
    logger.info(f"✅ 已将 LibreOffice `program` 目录添加到 sys.path: {LIBREOFFICE_PROGRAM_PATH}")

try:
    import uno
    from com.sun.star.beans import PropertyValue
    logger.info("✅ UNO 模块导入成功！")
except ImportError as e:
    logger.error("❌ UNO 模块导入失败，请检查 LibreOffice 是否正确安装:", e)
    sys.exit(1)

# ============ 启动 LibreOffice headless 服务 =============
libreoffice_proc = None

def start_libreoffice():
    """ 启动 LibreOffice headless 服务 """
    global libreoffice_proc
    if libreoffice_proc is None:
        libreoffice_path = app.config["LIBREOFFICE_PATH"]
        cmd = [
            libreoffice_path,
            "--headless",
            '--accept=socket,host=localhost,port=2002;urp;',
            "--norestore",
            "--nolockcheck",
            "--nodefault"
        ]
        try:
            libreoffice_proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            time.sleep(3)  # 等待 LibreOffice 启动
            logger.info("✅ LibreOfficePortable headless 服务已启动")
        except Exception as e:
            logger.error("❌ 启动 LibreOfficePortable 失败: %s", e)
            libreoffice_proc = None

def stop_libreoffice():
    """ 关闭 LibreOffice headless 服务 """
    global libreoffice_proc
    if libreoffice_proc:
        try:
            libreoffice_proc.terminate()
            libreoffice_proc.wait()
            logger.info("✅ LibreOfficePortable headless 服务已关闭")
        except Exception as e:
            logger.error("❌ 关闭 LibreOfficePortable 失败: %s", e)
        finally:
            libreoffice_proc = None

# 在 Flask 第一次收到请求前，确保 LibreOffice 运行
@app.before_request
def ensure_libreoffice_running():
    logger.info("🔄 确保 LibreOffice headless 服务正在运行...")
    start_libreoffice()

# 在应用启动时启动 LibreOffice
start_libreoffice()

# 在应用退出时终止 LibreOffice
atexit.register(stop_libreoffice)

@app.before_request
def log_database_setup():
    logger.info('Starting database setup...')
