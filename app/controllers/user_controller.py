from flask_jwt_extended import jwt_required
from app import db
from app.exceptions.exceptions import CustomAPIException
from app.models.result import ResponseTemplate
from app.models.role import Role
from app.models.user import User
from app.middleware.authorization import role_required

# 获取所有用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def get_users():
    try:
        users = User.query.all()
        return ResponseTemplate.success(
            data=[user.to_dict() for user in users],
            message="User list retrieved successfully"
        )
    except Exception as e:
        raise CustomAPIException("Material not found in the project", 404)


# 获取单个用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def get_user(user_id):
    try:
        user = User.query.get(user_id)
        if not user:
            raise CustomAPIException("Material not found in the project", 404)
        return ResponseTemplate.success(data=user.to_dict(), message="User details retrieved successfully")
    except Exception as e:
        raise CustomAPIException("Material not found in the project", 404)


# 创建新用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def create_user(data):
    try:
        if 'username' not in data or 'password' not in data or 'user_fullname' not in data:
            raise CustomAPIException("Material not found in the project", 404)

        # 检查用户名是否已存在
        if User.query.filter_by(username=data['username']).first():
            raise CustomAPIException("Material not found in the project", 404)

        # 创建用户
        new_user = User(
            username=data['username'],
            user_fullname=data['user_fullname'],
        )
        new_user.set_password(data['password'])  # 加密密码

        # 分配角色（可选）
        if 'roles' in data:
            roles = Role.query.filter(Role.name.in_(data['roles'])).all()
            new_user.roles.extend(roles)

        db.session.add(new_user)
        db.session.commit()
        return ResponseTemplate.success(message="User created successfully")

    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)


# 更新用户信息（仅管理员可操作）
@jwt_required()
@role_required("admin")
def update_user(user_id, data):
    try:
        user = User.query.get(user_id)
        if not user:
            return ResponseTemplate.error(message="User not found")

        # 允许更新的字段
        if 'user_fullname' in data:
            user.user_fullname = data['user_fullname']
        if 'username' in data:
            user.username = data['username']
        if 'password' in data:
            user.set_password(data['password'])

        # 更新角色
        if 'roles' in data:
            roles = Role.query.filter(Role.name.in_(data['roles'])).all()
            user.roles = roles

        db.session.commit()
        return ResponseTemplate.success(message="User updated successfully")

    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)


# 删除用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def delete_user(user_id):
    try:
        user = User.query.get(user_id)
        if not user:
            raise CustomAPIException("Material not found in the project", 404)

        db.session.delete(user)
        db.session.commit()
        return ResponseTemplate.success(message="User deleted successfully")

    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)


# 重置密码（仅管理员可操作）
@jwt_required()
@role_required("admin")
def reset_password(user_id):
    try:
        user = User.query.get(user_id)
        if not user:
            raise CustomAPIException("Material not found in the project", 404)

        new_password = user.reset_password()  # 重置密码
        db.session.commit()

        return ResponseTemplate.success(message=f"Password reset successfully. New password: {new_password}")

    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message=f"Error resetting password: {str(e)}")


# 禁用用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def disable_user(user_id):
    try:
        user = User.query.get(user_id)
        if not user:
            raise CustomAPIException("Material not found in the project", 404)

        user.status = "disabled"
        db.session.commit()
        return ResponseTemplate.success(message="User disabled successfully")

    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)


# 启用用户（仅管理员可操作）
@jwt_required()
@role_required("admin")
def enable_user(user_id):
    try:
        user = User.query.get(user_id)
        if not user:
            raise CustomAPIException("Material not found in the project", 404)

        user.status = "active"
        db.session.commit()
        return ResponseTemplate.success(message="User enabled successfully")

    except Exception as e:
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)
