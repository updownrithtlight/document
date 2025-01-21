from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import FieldOption
from app.exceptions.exceptions import CustomAPIException


@jwt_required()
def get_field_option_list():
    field_options = FieldOption.query.all()  # 查询所有字段选项
    field_option_list = [field_option.to_dict() for field_option in field_options]
    return ResponseTemplate.success(data=field_option_list, message='success')


@jwt_required()
def get_field_option(field_option_id):
    field_option = FieldOption.query.get(field_option_id)  # 根据 ID 获取字段选项
    if not field_option:
        raise CustomAPIException("FieldDefinition not found", 404)
    return ResponseTemplate.success(data=field_option.to_dict(), message='success')


@jwt_required()
def create_field_option(data):
    new_field_option = FieldOption(
        field_id=data['field_id'],
        parent_id=data.get('parent_id'),  # 父 ID 可选
        option_value=data['option_value'],
        image_path=data.get('image_path')  # 图片路径可选
    )
    db.session.add(new_field_option)
    db.session.commit()
    return ResponseTemplate.success(message='FieldOption created successfully')


@jwt_required()
def update_field_option(field_option_id, data):
    field_option = FieldOption.query.get(field_option_id)
    if not field_option:
        raise CustomAPIException("FieldDefinition not found", 404)

    field_option.field_id = data['field_id']
    field_option.parent_id = data.get('parent_id')  # 更新父 ID
    field_option.option_value = data['option_value']
    field_option.image_path = data.get('image_path')  # 更新图片路径
    db.session.commit()

    return ResponseTemplate.success(message='FieldOption updated successfully')


@jwt_required()
def delete_field_option(field_option_id):
    field_option = FieldOption.query.get(field_option_id)
    if not field_option:
        raise CustomAPIException("FieldDefinition not found", 404)
    db.session.delete(field_option)
    db.session.commit()
    return ResponseTemplate.success(message='FieldOption deleted successfully')
