from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import  ProjectFeature



@jwt_required()
def get_project_features(project_id):
    features = ProjectFeature.query.filter_by(project_id=project_id).order_by(ProjectFeature.sort_order).all()
    result = [{"feature_id": f.feature.id, "label": f.feature.label, "sort_order": f.sort_order} for f in features]
    return ResponseTemplate.success(data=result,message="获取成功！")


# **保存项目的技术特点**
@jwt_required()
def save_project_features(project_id,data):
    # 先删除该项目已有的技术特点
    ProjectFeature.query.filter_by(project_id=project_id).delete()

    # 插入新数据
    for feature_data in data:
        new_feature = ProjectFeature(
            project_id=project_id,
            feature_id=feature_data["feature_id"],
            sort_order=feature_data["sort_order"]
        )
        db.session.add(new_feature)

    db.session.commit()
    return ResponseTemplate.success(data=None,message="保存成功！")


def get_features(project_id):
    """
    查询项目下的特性列表，并返回可直接用于模板渲染的数据结构。

    :param project_id: 项目 ID
    :return: JSON 响应，包含一个 'data' 字段，内部含 'features' 列表，可直接传给模板
    """
    features = (ProjectFeature.query
                .filter_by(project_id=project_id)
                .order_by(ProjectFeature.sort_order)
                .all())

    # 构造数据列表，每个条目包含模板可直接使用的字段
    feature_list = []
    for f in features:
        feature_list.append({
            "feature_id": f.feature.id,
            "label": f.feature.label,
            "sort_order": f.sort_order
        })

    # 将 feature_list 包装为可直接在模板中使用的上下文
    # 如 docxtpl 中可以 {{ features }} 循环，也可以用 features.label
    result = {
        "features": feature_list
    }

    return result
