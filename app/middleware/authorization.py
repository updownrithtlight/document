from flask_jwt_extended import get_jwt_identity, verify_jwt_in_request
from functools import wraps

from app.exceptions.exceptions import CustomAPIException
from app.models.user import User

def role_required(required_role):
    """ 检查用户是否拥有指定角色 """
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            verify_jwt_in_request()
            user_identity = get_jwt_identity()
            user = User.query.get(user_identity)
            if not user or not user.has_role(required_role):
                raise CustomAPIException("联系管理员", 404)
            return fn(*args, **kwargs)
        return wrapper
    return decorator
