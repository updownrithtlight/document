from flask import request
from app import db, jwt_required, logger
from app.exceptions.exceptions import CustomAPIException
from app.models.result import ResponseTemplate
from app.models.models import ProjectMaterial, MaterialInfo

@jwt_required()
def get_project_materials(project_id):
    """
    根据 project_id 查询材料列表，并关联 MaterialInfo（无需分页）。
    """
    try:
        # 查询 ProjectMaterial 并关联 MaterialInfo
        materials = (
            ProjectMaterial.query
            .filter_by(project_id=project_id)
            .join(MaterialInfo, ProjectMaterial.material_id == MaterialInfo.id)
            .add_columns(
                ProjectMaterial.id.label('project_material_id'),
                MaterialInfo.id.label('material_id'),
                MaterialInfo.material_code,
                MaterialInfo.material_name,
                MaterialInfo.model_specification,
                MaterialInfo.unit,

            )
            .all()
        )

        # 格式化结果
        material_list = [
            {
                'project_material_id': m.project_material_id,
                'material_id': m.material_id,
                'material_code': m.material_code,
                'material_name': m.material_name,
                'unit': m.unit,
                'model_specification': m.model_specification

            }
            for m in materials
        ]

        # 返回数据
        return ResponseTemplate.success(
            data={
                "materials": material_list
            },
            message='success'
        )
    except Exception as e:
        logger.error(f"Error fetching project materials: {e}")
        return ResponseTemplate.error(message='Failed to fetch materials')


@jwt_required()
def save_project_material():
    """
    保存材料到项目。
    """
    try:
        # 获取请求数据
        data = request.json
        project_id = data.get('project_id')
        material_id = data.get('material_id')

        if not project_id or not material_id:
            raise CustomAPIException("Project ID and Material ID are required", 404)


        # 检查是否已存在
        existing = ProjectMaterial.query.filter_by(project_id=project_id, material_id=material_id).first()
        if existing:
            return ResponseTemplate.error(message='Material already exists in the project')

        # 保存到数据库
        new_project_material = ProjectMaterial(project_id=project_id, material_id=material_id)
        db.session.add(new_project_material)
        db.session.commit()

        return ResponseTemplate.success(
            message='Material saved successfully',
            data=new_project_material.to_dict()
        )
    except Exception as e:
        logger.error(f"Error saving project material: {e}")
        db.session.rollback()
        return ResponseTemplate.error(message='Failed to save material')


@jwt_required()
def remove_project_material(material_id):
    """
    从项目中移除材料。
    """
    try:
        # 获取项目 ID（从查询参数中）
        project_id = request.args.get('project_id')

        if not project_id:
            return ResponseTemplate.error(message='Project ID is required')

        # 查询对应的 ProjectMaterial
        project_material = ProjectMaterial.query.filter_by(project_id=project_id, material_id=material_id).first()
        if not project_material:
            raise CustomAPIException("Material not found in the project", 404)


        # 删除记录
        db.session.delete(project_material)
        db.session.commit()

        return ResponseTemplate.success(
            message='Material removed successfully',
            data=project_material.to_dict()
        )
    except Exception as e:
        logger.error(f"Error removing project material: {e}")
        db.session.rollback()
        raise CustomAPIException("Material not found in the project", 404)



def get_project_materials_info(project_id):
    """
    根据 project_id 查询材料列表，并关联 MaterialInfo（无需分页）。
    """
    try:
        # 查询 ProjectMaterial 并关联 MaterialInfo
        materials = (
            ProjectMaterial.query
            .filter_by(project_id=project_id)
            .join(MaterialInfo, ProjectMaterial.material_id == MaterialInfo.id)
            .add_columns(
                ProjectMaterial.id.label('project_material_id'),
                MaterialInfo.id.label('material_id'),
                MaterialInfo.material_code,
                MaterialInfo.material_name,
                MaterialInfo.model_specification,
                MaterialInfo.unit
            )
            .all()
        )

        # 格式化结果
        material_list = [
            {
                'project_material_id': m.project_material_id,
                'material_id': m.material_id,
                'material_code': m.material_code,
                'material_name': m.material_name,
                'model_specification': m.model_specification,
                'unit': m.unit,
            }
            for m in materials
        ]

        # 返回材料列表
        return material_list
    except Exception as e:
        logger.error(f"Error fetching project materials: {e}")
        return []  # 返回空列表以避免中断
