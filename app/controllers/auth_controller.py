from app import bcrypt, db, jwt
from app.models.result import ResponseTemplate
from app.models.user import User
from flask import request, make_response, current_app
from flask_jwt_extended import create_access_token, jwt_required, get_jwt_identity
from datetime import timedelta, datetime
import jwt as pyjwt
from app.exceptions.exceptions import CustomAPIException


# 注册用户
def register():
    data = request.get_json()

    # 检查用户名是否已存在
    if User.query.filter_by(username=data['username']).first():
        raise CustomAPIException("Username already exists", 400)

    # 密码加密
    hashed_password = bcrypt.generate_password_hash(data['password']).decode('utf-8')

    # 创建新用户
    new_user = User(username=data['username'], user_fullname=data['user_fullname'], password=hashed_password)

    # 添加到数据库
    try:
        db.session.add(new_user)
        db.session.commit()
        return ResponseTemplate.success(message="User registered successfully")
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException(f"Database error: {str(e)}", 500)

import json

def login():
    data = request.get_json()

    # 查找用户
    user = User.query.filter_by(username=data['username']).first()
    if not user or not bcrypt.check_password_hash(user.password, data['password']):
        raise CustomAPIException("Invalid username or password", 401)

    user_id_str = str(user.user_id)
    roles = [role.name for role in user.roles] if user.roles else ["user"]

    # ✅ 解决 `Subject must be a string`：用 JSON 字符串存 `identity`
    access_token = create_access_token(identity=user_id_str,
                                       expires_delta=current_app.config['JWT_ACCESS_TOKEN_EXPIRES'])

    refresh_token = pyjwt.encode(
        {'user_id': user_id_str, 'roles': roles, 'exp': datetime.utcnow() + timedelta(days=30)},
        current_app.config['JWT_SECRET_KEY'],
        algorithm=current_app.config['JWT_ALGORITHM']
    )

    response = make_response(
        ResponseTemplate.success(message="Login successful", data={"access_token": access_token})
    )
    response.set_cookie(
        'refresh_token', refresh_token,
        max_age=current_app.config['JWT_REFRESH_TOKEN_EXPIRES'],
        httponly=True, secure=True, path='/'
    )

    return response




# 续期 Access Token
def refresh():
    # 从 cookies 获取 refresh_token
    refresh_token = request.cookies.get('refresh_token')
    if not refresh_token:
        raise CustomAPIException("Refresh token is missing", 401)

    # 解码 refresh_token，验证有效性
    try:
        decoded_token = pyjwt.decode(
            refresh_token,
            current_app.config['JWT_SECRET_KEY'],
            algorithms=[current_app.config['JWT_ALGORITHM']]
        )
    except pyjwt.ExpiredSignatureError:
        raise CustomAPIException("Refresh token has expired", 401)
    except pyjwt.InvalidTokenError:
        raise CustomAPIException("Invalid refresh token", 401)

    # 提取 user_id 和 roles
    user_id_str = decoded_token.get('user_id')
    roles = decoded_token.get('roles', ["user"])
    if not user_id_str:
        raise CustomAPIException("Invalid refresh token payload", 401)

    # 生成新的 access_token

    # ✅ 解决 `Subject must be a string`：用 JSON 字符串存 `identity`
    new_access_token = create_access_token(identity=user_id_str,
                                       expires_delta=current_app.config['JWT_ACCESS_TOKEN_EXPIRES'])

    return ResponseTemplate.success(message="Access token refreshed", data={"access_token": new_access_token})


# 获取当前用户信息
@jwt_required()
def get_user_info():
    user_identity = get_jwt_identity()
    return ResponseTemplate.success(
        message="User info retrieved successfully",
        data={"user_id": user_identity["user_id"], "roles": user_identity["roles"]}
    )


# ✅ **用户退出登录**
def logout():
    response = make_response(ResponseTemplate.success(message="Logout successful"))

    # **📌 清除 `refresh_token`**
    response.set_cookie('refresh_token', '', max_age=0, httponly=True, secure=True, path='/')

    return response