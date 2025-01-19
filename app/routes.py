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
    app.register_blueprint(file_bp)  # 注册文件管理蓝图
