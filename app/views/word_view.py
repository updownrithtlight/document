from flask import Blueprint, request
from app.controllers import word_controller
from app.controllers import word_product_spec_controller

word_bp = Blueprint('word', __name__, url_prefix='/api/word')


@word_bp.route('/tech_manual/<int:project_id>', methods=['GET'])
def generate_tech_manual(project_id):
    return word_controller.generate_tech_manual(project_id)  # 生成技术说明书


@word_bp.route('/product_spec/<int:project_id>', methods=['GET'])
def generate_product_spec(project_id):
    return word_product_spec_controller.generate_product_spec(project_id)  # 生成产品规范
