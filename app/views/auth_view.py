from flask import Blueprint, request
from app.controllers import auth_controller

auth_bp = Blueprint('auth', __name__, url_prefix='/api/auth')


@auth_bp.route('/register', methods=['POST'])
def register():
    return auth_controller.register()


@auth_bp.route('/login', methods=['POST'])
def login():
    return auth_controller.login()



@auth_bp.route('/refresh-token', methods=['POST'])
def refresh():
    return auth_controller.refresh()


@auth_bp.route('/logout', methods=['POST'])
def logout():
    return auth_controller.logout()


@auth_bp.route('/me', methods=['GET'])
def get_user_info():
    return auth_controller.get_user_info()

@auth_bp.route('/updatePassword', methods=['POST'])
def update_password():
    data = request.json
    return auth_controller.update_password(data)

