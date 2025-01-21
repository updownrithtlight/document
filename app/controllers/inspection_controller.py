from flask import request
from app import db
from app.models.inspection import ProjectInspection, InspectionItem
from app.models.result import ResponseTemplate
from flask_jwt_extended import jwt_required

# 获取所有标准检验项目
@jwt_required()
def get_inspection_items():
    items = InspectionItem.query.all()
    return ResponseTemplate.success(data=[item.to_dict() for item in items], message="获取所有检验项目成功")

# 获取项目绑定的检验数据
@jwt_required()
def get_project_inspections(project_id):
    inspections = ProjectInspection.query.filter_by(project_id=project_id).all()
    return ResponseTemplate.success(data=[i.to_dict() for i in inspections], message="获取项目检验数据成功")

# 保存 `project_id` 绑定的检验数据
@jwt_required()
def save_project_inspections(project_id):
    try:
        data = request.get_json()
        inspections = data.get("inspections", [])

        if not project_id:
            return ResponseTemplate.error(message="缺少项目 ID")

        # 先删除当前项目的所有检验数据
        ProjectInspection.query.filter_by(project_id=project_id).delete()

        # 添加新数据
        for item in inspections:
            new_inspection = ProjectInspection(
                project_id=project_id,
                item_key=item.get("key"),
                pcb=item.get("pcb", False),
                before_seal=item.get("beforeSeal", False),
                after_label=item.get("afterLabel", False),
                sample_plan=item.get("samplePlan", "")
            )
            db.session.add(new_inspection)

        db.session.commit()
        return ResponseTemplate.success(message="检验项目保存成功")

    except Exception as e:
        db.session.rollback()
        return ResponseTemplate.error(message=f"数据保存失败: {str(e)}")
