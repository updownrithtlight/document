from flask import Blueprint, request
from app.controllers import material_info_controller

material_info_bp = Blueprint('material_info', __name__, url_prefix='/api/material_info')


@material_info_bp.route('/material_info', methods=['GET'])
def get_material_info_list():
    return material_info_controller.get_material_info_list()


@material_info_bp.route('/material_info', methods=['POST'])
def create_material_info():
    return material_info_controller.create_material_info(request.json)


@material_info_bp.route('/material_info/<int:material_info_id>', methods=['GET'])
def get_material_info(material_info_id):
    return material_info_controller.get_material_info(material_info_id)


@material_info_bp.route('/material_info/<int:material_info_id>', methods=['PUT'])
def update_material_info(material_info_id):
    return material_info_controller.update_material_info(material_info_id, request.json)


@material_info_bp.route('material_info/<int:material_info_id>', methods=['DELETE'])
def delete_material_info(material_info_id):
    return material_info_controller.delete_material_info(material_info_id)


# **✅ 添加批量导入 API**
@material_info_bp.route('/material_info/import', methods=['POST'])
def import_material_info():
    return material_info_controller.import_materials()