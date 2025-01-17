from flask import Blueprint, request
from app.controllers import field_option_controller

field_option_bp = Blueprint('field_option', __name__, url_prefix='/api/field_option')


@field_option_bp.route('/field_option-list', methods=['GET'])
def get_field_option_list():
    return field_option_controller.get_field_option_list()


@field_option_bp.route('/create', methods=['POST'])
def create_field_option():
    return field_option_controller.create_field_option(request.json)


@field_option_bp.route('/<int:field_option_id>', methods=['GET'])
def get_field_option(field_option_id):
    return field_option_controller.get_field_option(field_option_id)


@field_option_bp.route('/<int:field_option_id>', methods=['PUT'])
def update_field_option(field_option_id):
    return field_option_controller.update_field_option(field_option_id, request.json)


@field_option_bp.route('/<int:field_option_id>', methods=['DELETE'])
def delete_field_option(field_option_id):
    return field_option_controller.delete_field_option(field_option_id)
