from app import db

class Menu(db.Model):
    __tablename__ = 't_menu_item'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now())
    created_by = db.Column(db.String(255), default=None)
    updated_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), onupdate=db.func.now())
    updated_by = db.Column(db.String(255), default=None)
    component = db.Column(db.String(255), nullable=False)
    icon = db.Column(db.String(255), default=None)
    name = db.Column(db.String(255), nullable=False)
    path = db.Column(db.String(255), nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey('t_menu_item.id'))

    parent_menu = db.relationship('Menu', remote_side=[id], backref='children', lazy='subquery')

    def to_dict(self):
        return {
            'id': self.id,
            'created_at': self.created_at,
            'created_by': self.created_by,
            'updated_at': self.updated_at,
            'updated_by': self.updated_by,
            'component': self.component,
            'icon': self.icon,
            'name': self.name,
            'path': self.path,
            'parent_id': self.parent_id,
        }
