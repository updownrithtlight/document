from app import bcrypt, db, jwt
from app.models.result import ResponseTemplate
from app.models.models import User
from flask import request, make_response, current_app
from flask_jwt_extended import create_access_token, jwt_required, get_jwt_identity, get_jwt
from datetime import timedelta, datetime
import jwt as pyjwt



# 注册函数
def register():
    try:
        data = request.get_json()

        # 检查用户名是否已经存在
        if User.query.filter_by(username=data['username']).first():
            return ResponseTemplate.error(message='Username already exists')

        # 密码加密
        hashed_password = bcrypt.generate_password_hash(data['password']).decode('utf-8')

        # 创建新用户
        new_user = User(username=data['username'], user_fullname=data['user_fullname'], password=hashed_password)

        # 添加到数据库
        db.session.add(new_user)
        db.session.commit()

        return ResponseTemplate.success(message='User registered successfully')

    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message='An error occurred while registering user')


def login():
    data = request.get_json()

    # 验证用户名和密码
    user = User.query.filter_by(username=data['username']).first()
    if user and bcrypt.check_password_hash(user.password, data['password']):
        user_id_str = str(user.user_id)


        # 生成 access token
        access_token = create_access_token(identity=user_id_str, expires_delta=current_app.config['JWT_ACCESS_TOKEN_EXPIRES'])

        # 生成 refresh token（长时间有效）
        refresh_token = pyjwt.encode(
            {'user_id': user_id_str, 'exp': datetime.utcnow() + timedelta(days=30)},  # 设置过期时间
            current_app.config['JWT_SECRET_KEY'],  # 使用签名密钥
            algorithm=current_app.config['JWT_ALGORITHM']  # 使用指定的签名算法
        )

        # 将 refresh token 设置为 Cookie（httpOnly 防止 JS 获取，Secure 在 HTTPS 下使用）
        response = make_response(
            ResponseTemplate.success(message='Login successful', data={'access_token': access_token})
        )
        response.set_cookie('refresh_token', refresh_token, max_age=current_app.config['JWT_REFRESH_TOKEN_EXPIRES'], httponly=True,
                            secure=True, path='/')

        return response
    else:
        return ResponseTemplate.error(message='Invalid username or password')




def refresh():
    try:
        # 从 cookies 中获取 refresh_token
        refresh_token = request.cookies.get('refresh_token')
        if not refresh_token:
            return ResponseTemplate.error(message='Refresh token is missing')

        # 解码 refresh_token，验证有效性
        decoded_token = pyjwt.decode(
            refresh_token,
            current_app.config['JWT_SECRET_KEY'],
            algorithms=[current_app.config['JWT_ALGORITHM']]
        )

        # 提取 user_id
        user_id_str = decoded_token.get('user_id')
        if not user_id_str:
            return ResponseTemplate.error(message='Invalid refresh token')

        # 生成新的 access_token
        new_access_token = create_access_token(identity=user_id_str, expires_delta=current_app.config['JWT_ACCESS_TOKEN_EXPIRES'])

        # 创建响应并返回新的 access_token
        response = make_response(
            ResponseTemplate.success(message='Access token refreshed', data={'access_token': new_access_token})
        )

        return response

    except pyjwt.ExpiredSignatureError:
        return ResponseTemplate.error(message='Refresh token has expired')
    except pyjwt.InvalidTokenError:
        return ResponseTemplate.error(message='Invalid refresh token')
    except Exception as e:
        return ResponseTemplate.error(message=f'Error refreshing token: {str(e)}')
