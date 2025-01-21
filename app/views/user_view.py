from flask import Blueprint, request
from app.controllers import user_controller

user_bp = Blueprint('user', __name__, url_prefix='/api/users')

@user_bp.route('', methods=['GET'])
def get_users():
    return user_controller.get_users()

@user_bp.route('/<int:user_id>', methods=['GET'])
def get_user(user_id):
    return user_controller.get_user(user_id)

@user_bp.route('', methods=['POST'])
def create_user():
    data = request.json
    return user_controller.create_user(data)

@user_bp.route('/<int:user_id>', methods=['PUT'])
def update_user(user_id):
    data = request.json
    return user_controller.update_user(user_id, data)

@user_bp.route('/<int:user_id>', methods=['DELETE'])
def delete_user(user_id):
    return user_controller.delete_user(user_id)


@user_bp.route('/<int:user_id>/reset-password', methods=['PUT'])
def reset_password(user_id):
    return user_controller.reset_password(user_id)

@user_bp.route('/<int:user_id>/disable', methods=['PUT'])
def disable_user(user_id):
    return user_controller.disable_user(user_id)

@user_bp.route('/<int:user_id>/enable', methods=['PUT'])
def enable_user(user_id):
    return user_controller.enable_user(user_id)


@user_bp.route('/<int:user_id>/roles', methods=['PUT'])
def update_user_roles(user_id):
    return user_controller.update_user_roles(user_id)