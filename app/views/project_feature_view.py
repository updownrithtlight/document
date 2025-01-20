from flask import Blueprint, request
from app.controllers import project_feature_controller

project_feature_bp = Blueprint('project_feature', __name__, url_prefix='/api/project_feature')


@project_feature_bp.route("/<int:project_id>/features", methods=["GET"])
def get_project_features(project_id):
    return project_feature_controller.get_project_features(project_id)

# **保存项目的技术特点**
@project_feature_bp.route("/<int:project_id>/features", methods=["POST"])
def save_project_features(project_id):
    data = request.json.get("features", [])
    return project_feature_controller.save_project_features(project_id,data)

