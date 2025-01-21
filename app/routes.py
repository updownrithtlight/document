from app.views.auth_view import auth_bp
from app.views.project_field_view import project_field_bp
from app.views.menu_view import menu_bp
from app.views.project_view import project_bp
from app.views.field_definition_view import field_definition_bp
from app.views.material_info_view import material_info_bp
from app.views.field_option_view import field_option_bp
from app.views.project_material_view import project_material_bp
from app.views.excel_view import excel_bp
from app.views.word_view import word_bp
from app.views.file_view import file_bp  # 新增文件管理路由
from app.views.project_feature_view import project_feature_bp  # 新增文件管理路由
from app.views.technical_feature_view import feature_bp  # 新增文件管理路由
from app.views.project_important_notes_view import project_important_notes_bp
from app.views.important_notes_view import important_notes_bp
from app.views.user_view import user_bp
from app.views.role_view import role_bp


def setup_routes(app):
    app.register_blueprint(auth_bp)
    app.register_blueprint(project_field_bp)
    app.register_blueprint(menu_bp)
    app.register_blueprint(project_bp)
    app.register_blueprint(field_definition_bp)
    app.register_blueprint(material_info_bp)
    app.register_blueprint(field_option_bp)
    app.register_blueprint(project_material_bp)
    app.register_blueprint(excel_bp)
    app.register_blueprint(word_bp)
    app.register_blueprint(feature_bp)
    app.register_blueprint(project_feature_bp)
    app.register_blueprint(important_notes_bp)
    app.register_blueprint(project_important_notes_bp)
    app.register_blueprint(file_bp)
    app.register_blueprint(role_bp)
    app.register_blueprint(user_bp)  # 注册用户 API

