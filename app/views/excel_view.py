from flask import Blueprint, request
from app.controllers import excel_controller

excel_bp = Blueprint('excel', __name__, url_prefix='/api/excel')



@excel_bp.route('/<int:project_id>', methods=['GET'])
def generate_excel(project_id):
    return excel_controller.generate_excel(project_id)

