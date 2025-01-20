from flask import Blueprint, request
from app.controllers import technical_feature_controller

feature_bp = Blueprint('technical_feature', __name__, url_prefix='/api/technical_feature')


@feature_bp.route("", methods=["GET"])
def get_all_features():
    return technical_feature_controller.get_all_features()



