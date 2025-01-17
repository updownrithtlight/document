from flask_jwt_extended import jwt_required
from app import db
from app.models.result import ResponseTemplate
from app.models.models import Menu

# 获取所有菜单项的列表
@jwt_required()
def get_menu_list():
    try:
        menus = Menu.query.all()  # 获取所有菜单项
        menu_list = [menu.to_dict() for menu in menus]  # 使用 to_dict() 转换为字典
        return ResponseTemplate.success(data=menu_list, message='Success')
    except Exception as e:
        return ResponseTemplate.error(message=f'Error getting menu list: {str(e)}')

# 获取单个菜单项的详细信息
@jwt_required()
def get_menu(menu_id):
    try:
        menu = Menu.query.get(menu_id)  # 根据菜单 ID 获取菜单
        if not menu:
            return ResponseTemplate.error(message='Menu not found')
        return ResponseTemplate.success(data=menu.to_dict(), message='Success')
    except Exception as e:
        return ResponseTemplate.error(message=f'Error getting menu: {str(e)}')

# 创建新的菜单项
@jwt_required()
def create_menu(data):
    try:
        # 确保数据字段完整
        if 'menu_name' not in data or 'path' not in data:
            return ResponseTemplate.error(message="Required fields are missing (menu_name, path)")

        # 创建新菜单项
        new_menu = Menu(
            name=data['menu_name'],  # 从请求中获取菜单名称
            path=data['path'],  # 从请求中获取路径
            component=data.get('component'),  # 如果有组件字段
            icon=data.get('icon'),  # 如果有图标字段
            parent_id=data.get('parent_id')  # 如果有父菜单
        )

        # 添加并提交到数据库
        db.session.add(new_menu)
        db.session.commit()

        return ResponseTemplate.success(message='Menu created successfully')
    except Exception as e:
        db.session.rollback()  # 出错时回滚数据库事务
        return ResponseTemplate.error(message=f'Error creating menu: {str(e)}')

# 更新现有菜单项
@jwt_required()
def update_menu(menu_id, data):
    try:
        if not data or 'menu_name' not in data:
            return ResponseTemplate.error(message='Menu name is required')

        # 查找菜单项
        menu = Menu.query.get(menu_id)
        if not menu:
            return ResponseTemplate.error(message='Menu not found')

        # 更新菜单项
        menu.name = data['menu_name']
        menu.path = data.get('path', menu.path)  # 如果没有传递 `path`，保持原值
        menu.component = data.get('component', menu.component)  # 更新 `component` 字段
        menu.icon = data.get('icon', menu.icon)  # 更新 `icon` 字段
        menu.parent_id = data.get('parent_id', menu.parent_id)  # 更新 `parent_id` 字段

        db.session.commit()  # 提交更改
        return ResponseTemplate.success(message='Menu updated successfully')
    except Exception as e:
        db.session.rollback()  # 出错时回滚数据库事务
        return ResponseTemplate.error(message=f'Error updating menu: {str(e)}')

# 删除菜单项
@jwt_required()
def delete_menu(menu_id):
    try:
        menu = Menu.query.get(menu_id)  # 根据菜单 ID 获取菜单
        if not menu:
            return ResponseTemplate.error(message='Menu not found')

        db.session.delete(menu)  # 删除菜单项
        db.session.commit()  # 提交更改
        return ResponseTemplate.success(message='Menu deleted successfully')
    except Exception as e:
        db.session.rollback()  # 出错时回滚数据库事务
        return ResponseTemplate.error(message=f'Error deleting menu: {str(e)}')
