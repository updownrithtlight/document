from flask import Blueprint
from app.controllers import inspection_controller


inspection_bp = Blueprint('inspection', __name__, url_prefix='/api/inspections')


@inspection_bp.route('/items', methods=['GET'])
def get_inspection_items():
    return inspection_controller.get_inspection_items()  # 生成技术说明书


@inspection_bp.route('/project/<int:project_id>/inspections', methods=['GET'])
def get_project_inspections(project_id):
    return inspection_controller.get_project_inspections(project_id)  # 生成产品规范


@inspection_bp.route("/project/<int:project_id>/inspections", methods=["POST"])
def save_project_inspections(project_id):
    return inspection_controller.save_project_inspections(project_id)