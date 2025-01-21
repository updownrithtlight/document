from flask_jwt_extended import jwt_required
from app import db
from app.exceptions.exceptions import CustomAPIException
from app.models.result import ResponseTemplate
from app.models.role import Role
from app.models.menu import Menu

# 获取所有角色列表
@jwt_required()
def get_role_list():
    try:
        roles = Role.query.all()
        role_list = [role.to_dict() for role in roles]
        return ResponseTemplate.success(data=role_list, message='Success')
    except Exception as e:
        raise CustomAPIException("Material not found in the project", 404)

# 获取单个角色的详细信息
@jwt_required()
def get_role(role_id):
    try:
        role = Role.query.get(role_id)
        if not role:
            return ResponseTemplate.error(message='Role not found')
        return ResponseTemplate.success(data=role.to_dict(), message='Success')
    except Exception as e:
        raise CustomAPIException("Material not found in the project", 404)

# 创建新角色
@jwt_required()
def create_role(data):
    try:
        if 'role_name' not in data:
            return ResponseTemplate.error(message="Role name is required")

        # 创建新角色
        new_role = Role(name=data['role_name'])
        db.session.add(new_role)
        db.session.commit()
        return ResponseTemplate.success(message='Role created successfully')
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)

# 删除角色
@jwt_required()
def delete_role(role_id):
    try:
        role = Role.query.get(role_id)
        if not role:
            raise CustomAPIException("Material not found in the project", 404)

        db.session.delete(role)
        db.session.commit()
        return ResponseTemplate.success(message='Role deleted successfully')
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)

# 给角色分配菜单
@jwt_required()
def assign_menu_to_role(data):
    try:
        role_id = data.get("role_id")
        menu_ids = data.get("menu_ids")

        role = Role.query.get(role_id)
        if not role:
            raise CustomAPIException("Material not found in the project", 404)

        menus = Menu.query.filter(Menu.id.in_(menu_ids)).all()
        if not menus:
            raise CustomAPIException("Material not found in the project", 404)

        role.menus = menus  # 直接替换菜单
        db.session.commit()
        return ResponseTemplate.success(message="Menus assigned successfully")
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)
