from app import db

# 角色 & 菜单的多对多关系表

t_role_menus = db.Table(
    't_role_menus',
    db.Column('role_id', db.Integer, db.ForeignKey('t_roles.id', ondelete='CASCADE'), primary_key=True),
    db.Column('menu_id', db.Integer, db.ForeignKey('t_menu_item.id', ondelete='CASCADE'), primary_key=True),
    info={'create_after': ['t_roles', 't_menu_item']}  # 手动定义创建顺序
)

# 用户 & 角色的多对多关系表
t_user_roles = db.Table(
    't_user_roles',
    db.Column('user_id', db.Integer, db.ForeignKey('t_users.user_id', ondelete='CASCADE'), primary_key=True),
    db.Column('role_id', db.Integer, db.ForeignKey('t_roles.id', ondelete='CASCADE'), primary_key=True)
)

class Role(db.Model):
    __tablename__ = 't_roles'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

    menus = db.relationship('Menu', secondary=t_role_menus, backref=db.backref('roles', lazy='dynamic'))

    def __repr__(self):
        return f"Role('{self.name}', id={self.id})"

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'menus': [{'id': menu.id, 'name': menu.name, 'path': menu.path} for menu in self.menus],
        }
