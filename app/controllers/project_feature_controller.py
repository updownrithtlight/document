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
