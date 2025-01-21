from flask import Blueprint, request
from app.controllers import role_controller

role_bp = Blueprint('role', __name__, url_prefix='/api/roles')


@role_bp.route('', methods=['GET'])
def get_role_list():
    return role_controller.get_role_list()


@role_bp.route('/<int:role_id>', methods=['GET'])
def get_role(role_id):
    return role_controller.get_role(role_id)


@role_bp.route('', methods=['POST'])
def create_role():
    data = request.json
    return role_controller.create_role(data)


@role_bp.route('/<int:role_id>', methods=['DELETE'])
def delete_role(role_id):
    return role_controller.delete_role(role_id)


@role_bp.route('/assign-menu', methods=['POST'])
def assign_menu_to_role():
    data = request.json
    return role_controller.assign_menu_to_role(data)
