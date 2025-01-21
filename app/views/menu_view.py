from flask import Blueprint, request
from app.controllers import menu_controller

menu_bp = Blueprint('menu', __name__, url_prefix='/api/menu')

@menu_bp.route('', methods=['GET'])
def get_menu_list():
    return menu_controller.get_menu_list()

@menu_bp.route('/<int:menu_id>', methods=['GET'])
def get_menu(menu_id):
    return menu_controller.get_menu(menu_id)

@menu_bp.route('', methods=['POST'])
def create_menu():
    data = request.json
    return menu_controller.create_menu(data)

@menu_bp.route('/<int:menu_id>', methods=['PUT'])
def update_menu(menu_id):
    data = request.json
    return menu_controller.update_menu(menu_id, data)

@menu_bp.route('/<int:menu_id>', methods=['DELETE'])
def delete_menu(menu_id):
    return menu_controller.delete_menu(menu_id)

@menu_bp.route('/user-menus', methods=['GET'])
def get_user_menu():
    return menu_controller.get_user_menu()
