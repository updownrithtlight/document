from flask import Blueprint, request
from app.controllers import field_definition_controller

field_definition_bp = Blueprint('field_definition', __name__, url_prefix='/api/field_definition')


@field_definition_bp.route('', methods=['GET'])
def get_field_definition_list():
    return field_definition_controller.get_field_definition_list()


@field_definition_bp.route('', methods=['POST'])
def create_field_definition():
    return field_definition_controller.create_field_definition(request.json)


@field_definition_bp.route('/<identifier>', methods=['GET'])
def get_field_definition(identifier):
    """ 兼容 ID（整数）和 Code（字符串）的查询 """
    if identifier.isdigit():
        return field_definition_controller.get_field_definition(int(identifier))
    else:
        return field_definition_controller.get_field_definition_by_code(identifier)



@field_definition_bp.route('/<int:field_definition_id>', methods=['PUT'])
def update_field_definition(field_definition_id):
    return field_definition_controller.update_field_definition(field_definition_id, request.json)


@field_definition_bp.route('/<int:field_definition_id>', methods=['DELETE'])
def delete_field_definition(field_definition_id):
    return field_definition_controller.delete_field_definition(field_definition_id)
