from app import db, logger
from app.models.result import ResponseTemplate
from app.models.models import FieldDefinition
from flask_jwt_extended import jwt_required
from app.exceptions.exceptions import CustomAPIException


@jwt_required()
def get_field_definition_list():
    """ 获取所有字段定义（包括层级关系） """
    field_definitions = FieldDefinition.query.filter(FieldDefinition.parent_id.is_(None)).all()
    field_definition_list = [field_definition.to_dict(include_children=True) for field_definition in field_definitions]
    return ResponseTemplate.success(data=field_definition_list)


@jwt_required()
def get_field_definition(field_definition_id):
    """ 根据 ID 获取字段定义 """
    field_definition = FieldDefinition.query.get(field_definition_id)
    if not field_definition:
        raise CustomAPIException("FieldDefinition not found", 404)
    return ResponseTemplate.success(data=field_definition.to_dict(include_children=True))


@jwt_required()
def get_field_definition_by_code(code):
    """ 根据 code 获取字段定义 """
    field_definition = FieldDefinition.query.filter_by(code=code).first()
    if not field_definition:
        raise CustomAPIException("FieldDefinition with this code not found", 404)
    return ResponseTemplate.success(data=field_definition.to_dict(include_children=True))


@jwt_required()
def create_field_definition(data):
    """ 创建字段定义 """
    if 'code' in data and data['code']:  # 确保 code 唯一
        existing = FieldDefinition.query.filter_by(code=data['code']).first()
        if existing:
            raise CustomAPIException("Code must be unique", 400)

    try:
        new_field_definition = FieldDefinition(
            parent_id=data.get('parent_id'),
            field_name=data['field_name'],
            code=data.get('code'),  # 允许 `NULL`
            field_type=data['field_type'],
            remarks=data.get('remarks')
        )
        db.session.add(new_field_definition)
        db.session.commit()
        return ResponseTemplate.success(message="FieldDefinition created successfully")
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException(f"Database error: {str(e)}", 500)


@jwt_required()
def update_field_definition(field_definition_id, data):
    """ 更新字段定义 """
    field_definition = FieldDefinition.query.get(field_definition_id)
    if not field_definition:
        raise CustomAPIException("FieldDefinition not found", 404)

    # 确保 `code` 唯一
    if 'code' in data and data['code']:
        existing = FieldDefinition.query.filter(FieldDefinition.code == data['code'], FieldDefinition.id != field_definition_id).first()
        if existing:
            raise CustomAPIException("Code must be unique", 400)

    try:
        field_definition.parent_id = data.get('parent_id')
        field_definition.field_name = data['field_name']
        field_definition.code = data.get('code')
        field_definition.field_type = data['field_type']
        field_definition.remarks = data.get('remarks')
        db.session.commit()

        return ResponseTemplate.success(message="FieldDefinition updated successfully")
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException(f"Database error: {str(e)}", 500)


@jwt_required()
def delete_field_definition(field_definition_id):
    """ 删除字段定义（支持级联删除子字段） """
    field_definition = FieldDefinition.query.get(field_definition_id)
    if not field_definition:
        raise CustomAPIException("FieldDefinition not found", 404)

    try:
        db.session.delete(field_definition)
        db.session.commit()
        return ResponseTemplate.success(message="FieldDefinition deleted successfully")
    except Exception as e:
        db.session.rollback()
        raise CustomAPIException(f"Database error: {str(e)}", 500)


def get_fields_by_code():
    codes = [
        'circuit_diagram',
        'insertion_loss',
        'weight',
        'environmental_characteristics'
    ]
    # 执行查询，获取符合条件的所有记录
    results = FieldDefinition.query.filter(FieldDefinition.code.in_(codes)).all()
    # 使用字典推导，创建以 code 为键，FieldDefinition 对象的字典表示为值的字典
    fields_by_code = {item.code: item.to_dict() for item in results}
    return fields_by_code
