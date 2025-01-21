from flask_jwt_extended import jwt_required, get_jwt_identity
from app import db
from app.models.result import ResponseTemplate
from app.models.menu import Menu
from app.models.user import User

# 获取所有菜单项
@jwt_required()
def get_menu_list():
    try:
        menus = Menu.query.all()
        menu_list = [menu.to_dict() for menu in menus]
        return ResponseTemplate.success(data=menu_list, message='Success')
    except Exception as e:
        return ResponseTemplate.error(message=f'Error getting menu list: {str(e)}')

# 获取单个菜单项
@jwt_required()
def get_menu(menu_id):
    try:
        menu = Menu.query.get(menu_id)
        if not menu:
            return ResponseTemplate.error(message='Menu not found')
        return ResponseTemplate.success(data=menu.to_dict(), message='Success')
    except Exception as e:
        return ResponseTemplate.error(message=f'Error getting menu: {str(e)}')

# 创建新菜单
@jwt_required()
def create_menu(data):
    try:
        if 'menu_name' not in data or 'path' not in data:
            return ResponseTemplate.error(message="Required fields are missing (menu_name, path)")

        new_menu = Menu(
            name=data['menu_name'],
            path=data['path'],
            component=data.get('component'),
            icon=data.get('icon'),
            parent_id=data.get('parent_id')
        )

        db.session.add(new_menu)
        db.session.commit()

        return ResponseTemplate.success(message='Menu created successfully')
    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message=f'Error creating menu: {str(e)}')

# 更新菜单
@jwt_required()
def update_menu(menu_id, data):
    try:
        if not data or 'menu_name' not in data:
            return ResponseTemplate.error(message='Menu name is required')

        menu = Menu.query.get(menu_id)
        if not menu:
            return ResponseTemplate.error(message='Menu not found')

        menu.name = data['menu_name']
        menu.path = data.get('path', menu.path)
        menu.component = data.get('component', menu.component)
        menu.icon = data.get('icon', menu.icon)
        menu.parent_id = data.get('parent_id', menu.parent_id)

        db.session.commit()
        return ResponseTemplate.success(message='Menu updated successfully')
    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message=f'Error updating menu: {str(e)}')

# 删除菜单
@jwt_required()
def delete_menu(menu_id):
    try:
        menu = Menu.query.get(menu_id)
        if not menu:
            return ResponseTemplate.error(message='Menu not found')

        db.session.delete(menu)
        db.session.commit()
        return ResponseTemplate.success(message='Menu deleted successfully')
    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message=f'Error deleting menu: {str(e)}')

# 获取当前用户角色对应的菜单
@jwt_required()
def get_user_menu():
    try:
        user_identity = get_jwt_identity()
        user = User.query.get(user_identity)  # 这里 user_identity 可能是 JSON，需要取 "user_id"

        if not user:
            return ResponseTemplate.error(message="User not found")

        if user.has_role("admin"):
            # 管理员获取所有菜单
            menus = Menu.query.order_by(Menu.id.asc()).all()  # 假设有 `order` 字段
        else:
            # 普通用户基于角色获取菜单
            menus = set()
            for role in user.roles:
                for menu in role.menus:
                    menus.add(menu)

            # 按 `order` 字段排序（如果没有 `order`，可以改成 `id`）
            menus = sorted(menus, key=lambda x: x.id if hasattr(x, 'id') else x.id)

        return ResponseTemplate.success(
            data=[menu.to_dict() for menu in menus],
            message="User menus retrieved successfully"
        )

    except Exception as e:
        return ResponseTemplate.error(message=f"Error fetching user menus: {str(e)}")
