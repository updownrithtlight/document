from flask import Blueprint, request
from app.controllers import project_material_controller

project_material_bp = Blueprint('project_material', __name__, url_prefix='/api/project-material')


@project_material_bp.route('/<int:project_id>', methods=['GET'])
def get_project_materials(project_id):
    return project_material_controller.get_project_materials(project_id)


@project_material_bp.route('', methods=['POST'])
def save_project_material():
    return project_material_controller.save_project_material()


@project_material_bp.route('/<int:material_id>', methods=['DELETE'])
def remove_project_material(material_id):
    return project_material_controller.remove_project_material(material_id)
