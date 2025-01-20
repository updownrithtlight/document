from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import  TechnicalFeature



@jwt_required()
def get_all_features():
    features = TechnicalFeature.query.all()
    return ResponseTemplate.success(data=[{"id": f.id, "label": f.label} for f in features],message="查询成功！")

