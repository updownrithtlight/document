from flask import Blueprint, request
from app.controllers import project_field_controller

project_field_bp = Blueprint('project_field', __name__, url_prefix='/api/project_field')


@project_field_bp.route('', methods=['GET'])
def get_project_field_list():
    return project_field_controller.get_project_field_list()


@project_field_bp.route('', methods=['POST'])
def create_or_update_project_field():
    """ 🔥 创建或更新 ProjectField 记录，保证 (`project_id`, `field_id`) 唯一 """
    return project_field_controller.create_or_update_project_field(request.json)



@project_field_bp.route('/<int:project_field_id>', methods=['GET'])
def get_project_field(project_field_id):
    return project_field_controller.get_project_field(project_field_id)



@project_field_bp.route('/<int:project_field_id>', methods=['DELETE'])
def delete_project_field(project_field_id):
    return project_field_controller.delete_project_field(project_field_id)



@project_field_bp.route('/project/<int:project_id>', methods=['GET'])
def get_project_fields_by_project_id(project_id):
    """ 根据 projectId 查询项目的所有字段 """
    return project_field_controller.get_project_fields_by_project_id(project_id)


@project_field_bp.route('/<int:project_id>/field/<int:field_id>', methods=['DELETE'])
def delete_project_field_by_project_and_field(project_id, field_id):
    """ 根据 `project_id` 和 `field_id` 删除 """
    return project_field_controller.delete_project_field(project_id, field_id)


@project_field_bp.route('/<int:project_id>/field/<int:field_id>', methods=['GET'])
def get_project_field_by_project_id_field_id(project_id, field_id):
    return project_field_controller.get_project_field_by_project_id_field_id(project_id, field_id)


@project_field_bp.route('/<int:field_value_id>', methods=['DELETE'])
def delete_project_field_by_id(field_value_id):
    """ 根据 `id` 删除 """
    return project_field_controller.delete_project_field_by_id(field_value_id)


@project_field_bp.route('/batch', methods=['POST'])
def batch_create_or_update_project_fields():
    """ 🔥 批量创建或更新 ProjectField 记录 """
    data_list = request.json.get('fields', [])  # 接收多个字段
    return project_field_controller.batch_create_or_update_project_fields(data_list)


@project_field_bp.route('/project/<int:project_id>/<int:parent_id>', methods=['GET'])
def get_project_fields_by_project_id_parent_id(project_id,parent_id):
    """ 根据 projectId 查询项目的所有字段 """
    return project_field_controller.get_project_fields_by_project_id_parent_id(project_id,parent_id)