from app import db
from sqlalchemy.dialects.mysql import ENUM



class FieldOption(db.Model):
    __tablename__ = 't_field_options'

    id = db.Column(db.Integer, primary_key=True)
    field_id = db.Column(db.Integer, nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey('t_field_options.id'), nullable=True)
    option_value = db.Column(db.String(255), nullable=False)
    image_path = db.Column(db.String(255), nullable=True)

    parent = db.relationship('FieldOption', remote_side=[id], backref='children')

    def __repr__(self):
        return f"FieldOption(id={self.id}, field_id={self.field_id}, option_value='{self.option_value}')"

    def to_dict(self):
        return {
            'id': self.id,
            'field_id': self.field_id,
            'parent_id': self.parent_id,
            'option_value': self.option_value,
            'image_path': self.image_path,
        }


class MaterialInfo(db.Model):
    __tablename__ = 't_material_info'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    material_code = db.Column(db.String(50), nullable=False)
    material_name = db.Column(db.String(255), nullable=False)
    model_specification = db.Column(db.String(255), default=None)
    unit = db.Column(db.String(50), default=None)
    created_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp())
    updated_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp(), onupdate=db.func.current_timestamp())

    def __repr__(self):
        return f"MaterialInfo('{self.material_code}', '{self.material_name}')"

    def to_dict(self):
        return {
            'id': self.id,
            'material_code': self.material_code,
            'material_name': self.material_name,
            'model_specification': self.model_specification,
            'unit': self.unit,
            'created_at': self.created_at,
            'updated_at': self.updated_at,
        }


class ProjectFieldValue(db.Model):
    __tablename__ = 't_project_field_values'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_id = db.Column(db.Integer, db.ForeignKey('t_projects.id', ondelete='CASCADE'), nullable=False)
    field_id = db.Column(db.Integer, db.ForeignKey('t_field_definitions.id', ondelete='CASCADE'), nullable=False)
    is_checked = db.Column(db.Boolean, nullable=True, default=None)  # ✅ 复选框状态
    min_value = db.Column(db.String(255), nullable=True)
    typical_value = db.Column(db.String(255), nullable=True)
    max_value = db.Column(db.String(255), nullable=True)
    unit = db.Column(db.String(50), nullable=True)
    custom_value = db.Column(db.Text, nullable=True)  # ✅ 确保 `custom_value` 存在
    image_path = db.Column(db.String(500), nullable=True)
    description = db.Column(db.Text, nullable=True)
    code = db.Column(db.String(100),  nullable=True)  # 速记码，唯一
    parent_id = db.Column(db.Integer,nullable=True)
    product_code = db.Column(db.String(255), nullable=True)
    quantity = db.Column(db.String(255), nullable=True)
    remarks = db.Column(db.Text, nullable=True)

    def __repr__(self):
        return f"ProjectFieldValue(project_id={self.project_id}, field_id={self.field_id})"

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'field_id': self.field_id,
            'is_checked': self.is_checked,
            'min_value': self.min_value,
            'typical_value': self.typical_value,
            'max_value': self.max_value,
            'unit': self.unit,
            'custom_value': self.custom_value,  # ✅ 需要包含 `custom_value`
            'image_path': self.image_path,
            'description': self.description,
            'code': self.code,
            'product_code': self.product_code,
            'quantity': self.quantity,
            'remarks': self.remarks,
            'parent_id': self.parent_id,
        }



class Project(db.Model):
    __tablename__ = 't_projects'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_model = db.Column(db.String(50), nullable=False)
    project_name = db.Column(db.String(255), nullable=False)
    project_type = db.Column(db.String(100), nullable=True)
    working_temperature = db.Column(db.String(50), nullable=True, comment='工作温度')
    storage_temperature = db.Column(db.String(50), nullable=True, comment='存储温度')
    file_number = db.Column(db.String(100), nullable=False, comment='文件编号')
    product_number = db.Column(db.String(100), nullable=False, comment='产品编号')
    project_level = db.Column(ENUM('J', 'T', 'K', 'S'), nullable=False)
    created_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp())
    updated_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp(), onupdate=db.func.current_timestamp())

    def __repr__(self):
        return f"Project('{self.project_model}', '{self.project_name}', '{self.file_number}', '{self.product_number}')"

    def to_dict(self):
        return {
            'id': self.id,
            'project_model': self.project_model,
            'project_name': self.project_name,
            'project_type': self.project_type,
            'project_level': self.project_level,
            'working_temperature': self.working_temperature,
            'storage_temperature': self.storage_temperature,
            'file_number': self.file_number,
            'product_number': self.product_number,
            'created_at': self.created_at,
            'updated_at': self.updated_at,
        }


class FieldDefinition(db.Model):
    __tablename__ = 't_field_definitions'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    parent_id = db.Column(db.Integer, db.ForeignKey('t_field_definitions.id', ondelete='CASCADE'), nullable=True)
    field_name = db.Column(db.String(255), nullable=False)
    code = db.Column(db.String(100), unique=True, nullable=True)  # 速记码，唯一
    field_type = db.Column(db.Enum('select', 'input', 'image', 'checkbox', 'group', 'select+input'), nullable=False)
    remarks = db.Column(db.Text, nullable=True)

    # 关系映射（支持父子结构）
    parent = db.relationship("FieldDefinition", remote_side=[id], backref="children", lazy="joined")

    def __repr__(self):
        return f"FieldDefinition('{self.field_name}', '{self.code}', '{self.field_type}')"

    def to_dict(self, include_children=False):
        """ 转换为 JSON 格式，可选包含子字段 """
        data = {
            'id': self.id,
            'parent_id': self.parent_id,
            'field_name': self.field_name,
            'code': self.code,
            'field_type': self.field_type,
            'remarks': self.remarks,
        }
        if include_children:
            data['children'] = [child.to_dict() for child in self.children]
        return data



class ProjectMaterial(db.Model):
    __tablename__ = 't_project_material'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_id = db.Column(db.Integer, db.ForeignKey('t_projects.id', ondelete='CASCADE'), nullable=False)
    material_id = db.Column(db.Integer, db.ForeignKey('t_material_info.id', ondelete='CASCADE'), nullable=False)
    created_at = db.Column(db.TIMESTAMP, nullable=False, server_default=db.func.current_timestamp())

    # 关系映射
    project = db.relationship('Project', backref='project_materials', lazy='joined')
    material = db.relationship('MaterialInfo', backref='material_projects', lazy='joined')

    def __repr__(self):
        return f"ProjectMaterial(project_id={self.project_id}, material_id={self.material_id})"

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'material_id': self.material_id,
            'created_at': self.created_at,
        }



class TechnicalFeature(db.Model):
    """ 技术特点数据表（数据源） """
    __tablename__ = "t_technical_features"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    label = db.Column(db.String(255), nullable=False, comment="技术特点描述")

    def __repr__(self):
        return f"TechnicalFeature(id={self.id}, label='{self.label}')"

    def to_dict(self):
        """ 转换为 JSON 格式 """
        return {
            "id": self.id,
            "label": self.label,
        }


class ProjectFeature(db.Model):
    """ 关联项目与技术特点，并存储排序 """
    __tablename__ = "t_project_features"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_id = db.Column(db.Integer, db.ForeignKey("t_projects.id", ondelete="CASCADE"), nullable=False, comment="项目 ID")
    feature_id = db.Column(db.Integer, db.ForeignKey("t_technical_features.id", ondelete="CASCADE"), nullable=False, comment="技术特点 ID")
    sort_order = db.Column(db.Integer, nullable=False, comment="排序顺序")

    # 关系映射
    project = db.relationship("Project", backref=db.backref("project_features", cascade="all, delete-orphan", lazy="joined"))
    feature = db.relationship("TechnicalFeature", backref=db.backref("feature_projects", cascade="all, delete-orphan", lazy="joined"))

    def __repr__(self):
        return f"ProjectFeature(project_id={self.project_id}, feature_id={self.feature_id}, sort_order={self.sort_order})"

    def to_dict(self):
        """ 转换为 JSON 格式 """
        return {
            "id": self.id,
            "project_id": self.project_id,
            "feature_id": self.feature_id,
            "sort_order": self.sort_order,
            "feature_label": self.feature.label if self.feature else None  # 关联技术特点名称
        }


class ImportantNote(db.Model):
    """ 技术特点数据表（数据源） """
    __tablename__ = "t_important_notes"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    label = db.Column(db.String(255), nullable=False, comment="技术特点描述")

    def __repr__(self):
        return f"ImportantNote(id={self.id}, label='{self.label}')"

    def to_dict(self):
        """ 转换为 JSON 格式 """
        return {
            "id": self.id,
            "label": self.label,
        }


class ProjectImportantNote(db.Model):
    """ 关联项目与技术特点，并存储排序 """
    __tablename__ = "t_project_important_notes"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_id = db.Column(db.Integer, db.ForeignKey("t_projects.id", ondelete="CASCADE"), nullable=False, comment="项目 ID")
    note_id = db.Column(db.Integer, db.ForeignKey("t_important_notes.id", ondelete="CASCADE"), nullable=False, comment="技术特点 ID")
    sort_order = db.Column(db.Integer, nullable=False, comment="排序顺序")

    # 关系映射
    project = db.relationship("Project", backref=db.backref("project_important_notes", cascade="all, delete-orphan", lazy="joined"))
    note = db.relationship("ImportantNote", backref=db.backref("note_projects", cascade="all, delete-orphan", lazy="joined"))

    def __repr__(self):
        return f"ProjectImportantNote(project_id={self.project_id}, note_id={self.note_id}, sort_order={self.sort_order})"

    def to_dict(self):
        """ 转换为 JSON 格式 """
        return {
            "id": self.id,
            "project_id": self.project_id,
            "note_id": self.note_id,
            "sort_order": self.sort_order,
            "note_label": self.note.label if self.note else None  # 关联技术特点名称
        }


