from app import db, jwt_required, logger
from app.exceptions.exceptions import CustomAPIException
from app.models.result import ResponseTemplate
from app.models.models import ProjectFieldValue
import json
from app.controllers.field_definition_controller import get_field_name_by_code


@jwt_required()
def get_project_field_list():
    """ 获取所有项目字段 """
    project_fields = ProjectFieldValue.query.all()
    project_field_list = [project_field.to_dict() for project_field in project_fields]
    return ResponseTemplate.success(data=project_field_list, message='success')


@jwt_required()
def get_project_field(project_field_id):
    """ 根据 ID 获取单个项目字段 """
    project_field = ProjectFieldValue.query.get(project_field_id)
    if not project_field:
        return ResponseTemplate.success(message='ProjectFieldValue not found')
    return ResponseTemplate.success(data=project_field.to_dict(), message='success')


@jwt_required()
def get_project_fields_by_project_id(project_id):
    """ 根据 project_id 获取所有关联的 ProjectFieldValue 记录 """
    project_fields = ProjectFieldValue.query.filter_by(project_id=project_id).all()
    project_field_list = [project_field.to_dict() for project_field in project_fields]
    return ResponseTemplate.success(data=project_field_list, message='success')


@jwt_required()
def create_or_update_project_field(data):
    """ 🔥 增量更新项目字段，确保 (`project_id`, `field_id`) 唯一 🔥 """

    project_id = data.get('project_id')
    field_id = data.get('field_id')

    if not project_id or not field_id:
        return ResponseTemplate.error(message="Missing `project_id` or `field_id`", status=400)

    # ✅ 处理 custom_value，确保其为 JSON 字符串
    if "custom_value" in data:
        if isinstance(data["custom_value"], (list, dict)):
            data["custom_value"] = json.dumps(data["custom_value"])  # 转换为 JSON 字符串
        elif data["custom_value"] in [None, ""]:
            data["custom_value"] = None  # 避免空值存入数据库

    # ✅ 检查是否已有记录
    existing_project_field = ProjectFieldValue.query.filter_by(project_id=project_id, field_id=field_id).first()

    updatable_fields = [
        "is_checked", "min_value", "typical_value", "max_value", "unit",
        "custom_value", "image_path", "description", "parent_id", "code",
        "product_code", "quantity", "remarks"
    ]

    if existing_project_field:
        # ✅ **已存在，则增量更新**
        for field in updatable_fields:
            if field in data and data[field] is not None:  # 只更新传入且非 None 的字段
                setattr(existing_project_field, field, data[field])

        db.session.commit()
        return ResponseTemplate.success(message="ProjectFieldValue updated successfully")

    else:
        # ✅ **如果不存在，则插入新数据**
        new_project_field = ProjectFieldValue(
            project_id=project_id,
            field_id=field_id,
            **{field: data[field] for field in updatable_fields if field in data and data[field] is not None}
        )
        db.session.add(new_project_field)
        db.session.commit()
        return ResponseTemplate.success(message="ProjectFieldValue created successfully")

@jwt_required()
def delete_project_field(project_id, field_id):
    """ 🔥 根据 (`project_id`, `field_id`) 删除项目字段记录 """
    project_field = ProjectFieldValue.query.filter_by(project_id=project_id, field_id=field_id).first()
    if not project_field:
        raise CustomAPIException("ProjectFieldValue not found", 404)


    db.session.delete(project_field)
    db.session.commit()
    return ResponseTemplate.success(message='ProjectFieldValue deleted successfully')

@jwt_required()
def delete_project_field_by_id(field_value_id):
    """ 🔥 根据 `id` 删除项目字段记录 """
    project_field = ProjectFieldValue.query.get(field_value_id)
    if not project_field:
        raise CustomAPIException("ProjectFieldValue not found", 500)


    db.session.delete(project_field)
    db.session.commit()
    return ResponseTemplate.success(message='ProjectFieldValue deleted successfully')


@jwt_required()
def batch_create_or_update_project_fields(data_list):
    """ 🔥 批量 `UPSERT` ProjectField 记录 """
    for data in data_list:
        create_or_update_project_field(data)  # 直接复用 `create_or_update_project_field()`
    return ResponseTemplate.success(message="Batch ProjectField update successful")

@jwt_required()
def get_project_fields_by_project_id_parent_id(project_id, parent_id):
    """ 根据 project_id 获取所有关联的 ProjectFieldValue 记录 """
    project_fields = ProjectFieldValue.query.filter_by(project_id=project_id,parent_id=parent_id).all()
    project_field_list = [project_field.to_dict() for project_field in project_fields]
    return ResponseTemplate.success(data=project_field_list, message='success')


@jwt_required()
def get_project_field_by_project_id_field_id(project_id, field_id):
    existing_project_field = ProjectFieldValue.query.filter_by(project_id=project_id, field_id=field_id).first()
    if not existing_project_field:
        return ResponseTemplate.success(message='ProjectFieldValue not found')
    return ResponseTemplate.success(data=existing_project_field.to_dict(), message='success')




def get_list_by_project_id(project_id):
    """ 根据 project_id 获取所有关联的 ProjectFieldValue 记录 """
    project_fields = ProjectFieldValue.query.filter_by(project_id=project_id).all()
    return [project_field.to_dict() for project_field in project_fields]


def get_fields_by_project_id_parent_id(project_id, parent_id):
    """
    根据 project_id 和 parent_id 查询 t_project_field_values 表，
    并返回一个二维列表，每个子列表包含 [code, product_code, quantity, remarks]。
    """
    project_fields = ProjectFieldValue.query.filter_by(
        project_id=project_id,
        parent_id=parent_id
    ).all()

    rows_to_add = []
    for pf in project_fields:
        # 如果任何字段可能为空，可使用 or "" 做兜底处理

        rows_to_add.append([
            get_field_name_by_code(pf.code) or "",
            pf.product_code or "",
            pf.quantity or "",
            pf.remarks or ""
        ])

    return rows_to_add
