from flask import Blueprint, request
from app.controllers import word_controller

word_bp = Blueprint('word', __name__, url_prefix='/api/word')



@word_bp.route('/<int:project_id>', methods=['GET'])
def generate_word(project_id):
    return word_controller.generate_word(project_id)

