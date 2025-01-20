from flask import Blueprint, request
from app.controllers import project_important_notes_controller

project_important_notes_bp = Blueprint('project_important_notes', __name__, url_prefix='/api/project_important_notes')


@project_important_notes_bp.route("/<int:project_id>/notes", methods=["GET"])
def get_project_important_notes(project_id):
    return project_important_notes_controller.get_project_important_notes(project_id)

# **保存项目的技术特点**
@project_important_notes_bp.route("/<int:project_id>/notes", methods=["POST"])
def save_project_important_notes(project_id):
    data = request.json.get("notes", [])
    return project_important_notes_controller.save_project_important_notes(project_id,data)

