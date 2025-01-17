from flask import Blueprint, request
from app.controllers import project_controller

project_bp = Blueprint('project', __name__, url_prefix='/api/project')


@project_bp.route('/project', methods=['GET'])
def get_project_list():
    return project_controller.get_project_list()


@project_bp.route('/project', methods=['POST'])
def create_project():
    return project_controller.create_project(request.json)


@project_bp.route('/project/<int:project_id>', methods=['GET'])
def get_project(project_id):
    return project_controller.get_project(project_id)


@project_bp.route('/project/<int:project_id>', methods=['PUT'])
def update_project(project_id):
    return project_controller.update_project(project_id, request.json)


@project_bp.route('project/<int:project_id>', methods=['DELETE'])
def delete_project(project_id):
    return project_controller.delete_project(project_id)
