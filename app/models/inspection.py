from app import db

class InspectionItem(db.Model):
    """全局检验项目定义表"""
    __tablename__ = 't_inspection_items'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    key = db.Column(db.String(50), unique=True, nullable=False, comment="唯一键")
    name = db.Column(db.String(255), nullable=False, comment="检验项目名称")

    def to_dict(self):
        return {
            "key": self.key,
            "name": self.name
        }


class ProjectInspection(db.Model):
    """项目检验项目信息"""
    __tablename__ = 't_project_inspections'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    project_id = db.Column(db.Integer, nullable=False, index=True, comment="项目ID")
    item_key = db.Column(db.String(50), db.ForeignKey('t_inspection_items.key', ondelete='CASCADE'), nullable=False, comment="检验项目的 key")
    pcb = db.Column(db.Boolean, default=False, comment="印制件")
    before_seal = db.Column(db.Boolean, default=False, comment="灌封前")
    after_label = db.Column(db.Boolean, default=False, comment="贴标后")
    sample_plan = db.Column(db.String(100), nullable=True, comment="抽样方案")

    item = db.relationship("InspectionItem", backref="project_inspections")

    def to_dict(self):
        return {
            "project_id": self.project_id,
            "key": self.item_key,
            "name": self.item.name,  # 获取项目名称
            "pcb": self.pcb,
            "beforeSeal": self.before_seal,
            "afterLabel": self.after_label,
            "samplePlan": self.sample_plan
        }
