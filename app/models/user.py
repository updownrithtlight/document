from app import db, bcrypt
from app.models.role import Role, t_user_roles

class User(db.Model):
    __tablename__ = 't_users'

    user_id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_fullname = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(255), nullable=False)
    username = db.Column(db.String(100), nullable=False, unique=True)
    status = db.Column(db.String(20), nullable=False, default="active")  # 新增状态字段
    created_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp())

    roles = db.relationship('Role', secondary=t_user_roles, backref=db.backref('users', lazy='dynamic'))

    def set_password(self, password):
        self.password = bcrypt.generate_password_hash(password).decode('utf-8')

    def check_password(self, password):
        return bcrypt.check_password_hash(self.password, password)

    def reset_password(self):
        """重置密码为 `用户名 + 123`"""
        new_password = self.username + "123"
        self.set_password(new_password)
        return new_password  # 返回新密码，方便管理员通知用户

    def has_role(self, role_name):
        """检查用户是否拥有指定角色"""
        return any(role.name == role_name for role in self.roles)

    def to_dict(self):
        return {
            'user_id': self.user_id,
            'user_fullname': self.user_fullname,
            'username': self.username,
            'roles': [role.name for role in self.roles],  # 获取角色名称列表
            'status': self.status,
            'created_at': self.created_at,
        }
