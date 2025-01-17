from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import ProjectFieldValue


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
        return ResponseTemplate.error(message='ProjectFieldValue not found')
    return ResponseTemplate.success(data=project_field.to_dict(), message='success')


@jwt_required()
def get_project_fields_by_project_id(project_id):
    """ 根据 project_id 获取所有关联的 ProjectFieldValue 记录 """
    project_fields = ProjectFieldValue.query.filter_by(project_id=project_id).all()
    project_field_list = [project_field.to_dict() for project_field in project_fields]
    return ResponseTemplate.success(data=project_field_list, message='success')


@jwt_required()
def create_or_update_project_field(data):
    """ 🔥 创建或更新项目字段记录，确保唯一 (`project_id`, `field_id`) 🔥 """

    project_id = data['project_id']
    field_id = data['field_id']
    is_checked = data.get('is_checked', False)  # ✅ 默认 `False`

    # ✅ 先检查是否已存在相同 (`project_id`, `field_id`) 组合
    existing_project_field = ProjectFieldValue.query.filter_by(project_id=project_id, field_id=field_id).first()

    if existing_project_field:
        # ✅ **已存在，则更新**
        existing_project_field.is_checked = is_checked
        existing_project_field.min_value = data.get('min_value')
        existing_project_field.typical_value = data.get('typical_value')
        existing_project_field.max_value = data.get('max_value')
        existing_project_field.unit = data.get('unit')
        existing_project_field.custom_value = data.get('custom_value')
        existing_project_field.image_path = data.get('image_path')
        existing_project_field.description = data.get('description')

        db.session.commit()
        return ResponseTemplate.success(message='ProjectFieldValue updated successfully')

    else:
        # ✅ **如果不存在，则插入新数据**
        new_project_field = ProjectFieldValue(
            project_id=project_id,
            field_id=field_id,
            is_checked=is_checked,
            min_value=data.get('min_value'),
            typical_value=data.get('typical_value'),
            max_value=data.get('max_value'),
            unit=data.get('unit'),
            custom_value=data.get('custom_value'),
            image_path=data.get('image_path'),
            description=data.get('description')
        )
        db.session.add(new_project_field)
        db.session.commit()
        return ResponseTemplate.success(message='ProjectFieldValue created successfully')


@jwt_required()
def delete_project_field(project_id, field_id):
    """ 🔥 根据 (`project_id`, `field_id`) 删除项目字段记录 """
    project_field = ProjectFieldValue.query.filter_by(project_id=project_id, field_id=field_id).first()
    if not project_field:
        return ResponseTemplate.error(message='ProjectFieldValue not found')

    db.session.delete(project_field)
    db.session.commit()
    return ResponseTemplate.success(message='ProjectFieldValue deleted successfully')

@jwt_required()
def delete_project_field_by_id(field_value_id):
    """ 🔥 根据 `id` 删除项目字段记录 """
    project_field = ProjectFieldValue.query.get(field_value_id)
    if not project_field:
        return ResponseTemplate.error(message='ProjectFieldValue not found')

    db.session.delete(project_field)
    db.session.commit()
    return ResponseTemplate.success(message='ProjectFieldValue deleted successfully')


@jwt_required()
def batch_create_or_update_project_fields(data_list):
    """ 🔥 批量 `UPSERT` ProjectField 记录 """
    for data in data_list:
        create_or_update_project_field(data)  # 直接复用 `create_or_update_project_field()`
    return ResponseTemplate.success(message="Batch ProjectField update successful")
