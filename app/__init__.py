from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_jwt_extended import JWTManager, jwt_required
from flask_cors import CORS
import logging
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


from app.routes import setup_routes
setup_routes(app)

# 捕获 SQLAlchemy 的日志，并输出 SQL 语句
logging.getLogger('sqlalchemy.engine').setLevel(logging.INFO)

logging.basicConfig(level=logging.DEBUG)
app.logger.addHandler(logging.StreamHandler())


# 在应用启动时输出日志
@app.before_request
def log_database_setup():
    logger.info('Starting database setup...')



