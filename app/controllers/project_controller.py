from flask import request
from app import db, logger
from flask_jwt_extended import jwt_required
from app.models.result import ResponseTemplate
from app.models.models import Project, ProjectMaterial


@jwt_required()
def get_project_list():
    """分页查询和组合搜索项目列表"""
    # 获取分页参数
    page = request.args.get('page', 1, type=int)  # 当前页码，默认为第 1 页
    per_page = request.args.get('pageSize', 10, type=int)  # 每页显示的记录数，默认为 10 条

    # 获取搜索参数
    project_model = request.args.get('project_model', '').strip()
    project_name = request.args.get('project_name', '').strip()
    project_type = request.args.get('project_type', '').strip()
    project_level = request.args.get('project_level', '').strip()
    working_temperature = request.args.get('working_temperature', '').strip()
    storage_temperature = request.args.get('storage_temperature', '').strip()
    file_number = request.args.get('file_number', '').strip()
    product_number = request.args.get('product_number', '').strip()

    # 构建查询
    query = Project.query

    # 动态添加过滤条件
    if project_model:
        query = query.filter(Project.project_model.ilike(f"%{project_model}%"))
    if project_level:
        query = query.filter(Project.project_level == project_level)
    if project_name:
        query = query.filter(Project.project_name.ilike(f"%{project_name}%"))
    if project_type:
        query = query.filter(Project.project_type.ilike(f"%{project_type}%"))
    if working_temperature:
        query = query.filter(Project.working_temperature.ilike(f"%{working_temperature}%"))
    if storage_temperature:
        query = query.filter(Project.storage_temperature.ilike(f"%{storage_temperature}%"))
    if file_number:
        query = query.filter(Project.file_number.ilike(f"%{file_number}%"))
    if product_number:
        query = query.filter(Project.product_number.ilike(f"%{product_number}%"))

    # 查询数据库并分页
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)
    projects = pagination.items  # 当前页的项目列表

    # 转换为字典格式
    project_list = [project.to_dict() for project in projects]

    # 返回分页信息
    return ResponseTemplate.success(
        data={
            'projects': project_list,
            'totalElements': pagination.total,  # 总记录数
            'pages': pagination.pages,  # 总页数
            'page': pagination.page,  # 当前页码
            'pageSize': pagination.per_page,  # 每页显示的记录数
        },
        message='success'
    )



@jwt_required()
def get_project(project_id):
    """根据项目 ID 获取单个项目"""
    project = Project.query.get(project_id)
    if not project:
        return ResponseTemplate.error(message='Project not found')
    return ResponseTemplate.success(data=project.to_dict(), message='success')


@jwt_required()
def create_project(data):
    """创建新项目"""
    new_project = Project(
        project_model=data['project_model'],
        project_name=data['project_name'],
        project_type=data.get('project_type'),
        project_level=data.get('project_level'),
        working_temperature=data.get('working_temperature'),
        storage_temperature=data.get('storage_temperature'),
        file_number=data['file_number'],
        product_number=data['product_number']
    )
    db.session.add(new_project)
    db.session.commit()

    return ResponseTemplate.success(
        message='Project created successfully',
        data={'projectId': new_project.id}
    )
@jwt_required()
def update_project(project_id, data):
    """更新项目"""
    project = Project.query.get(project_id)
    if not project:
        return ResponseTemplate.error(message='Project not found')

    project.project_model = data['project_model']
    project.project_name = data['project_name']
    project.project_type = data.get('project_type')

    # ✅ 修正 tuple 变 string
    project_level = data.get('project_level')
    if isinstance(project_level, (list, tuple)):
        project_level = project_level[0]  # 取第一个元素，确保是字符串
    project.project_level = str(project_level)  # 直接赋值字符串

    print('project_level:', project_level)  # Debug 确保是 "T" 而不是 ('T',)

    project.working_temperature = data.get('working_temperature')
    project.storage_temperature = data.get('storage_temperature')
    project.file_number = data['file_number']
    project.product_number = data['product_number']

    db.session.commit()
    return ResponseTemplate.success(message='Project updated successfully')

@jwt_required()
def delete_project(project_id):
    """手动删除级联数据，确保 project_id 不为空"""
    db.session.query(ProjectMaterial).filter(ProjectMaterial.project_id == project_id).delete()
    db.session.commit()

    project = Project.query.get(project_id)
    if not project:
        return ResponseTemplate.error(message='Project not found')

    db.session.delete(project)
    db.session.commit()
    return ResponseTemplate.success(message='Project deleted successfully')

