from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import ImportantNote


@jwt_required()
def get_all_important_notes():
    """ 获取所有技术特点（数据源） """
    notes = ImportantNote.query.all()
    result = [{"id": note.id, "label": note.label} for note in notes]
    return ResponseTemplate.success(data=result, message="查询成功！")
