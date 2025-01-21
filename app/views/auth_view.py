from flask import Blueprint
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

