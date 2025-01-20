from flask import Blueprint, request
from app.controllers import important_notes_controller

important_notes_bp = Blueprint('important_notes', __name__, url_prefix='/api/important_notes')


@important_notes_bp.route("", methods=["GET"])
def get_all_important_notes():
    return important_notes_controller.get_all_important_notes()



