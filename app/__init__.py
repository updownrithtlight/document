from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_jwt_extended import JWTManager, jwt_required
from flask_cors import CORS
import logging
from config import Config
import pytz

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

# 统一设置 Flask 解析时区
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

# 捕获 SQLAlchemy 的日志，并输出 SQL 语句
logging.getLogger('sqlalchemy.engine').setLevel(logging.INFO)

logging.basicConfig(level=logging.DEBUG)
app.logger.addHandler(logging.StreamHandler())


# 在应用启动时输出日志
@app.before_request
def log_database_setup():
    logger.info('Starting database setup...')



