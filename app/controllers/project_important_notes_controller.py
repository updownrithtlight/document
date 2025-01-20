from app import db, jwt_required, logger
from app.models.result import ResponseTemplate
from app.models.models import ProjectImportantNote





@jwt_required()
def get_project_important_notes(project_id):
    """ 获取项目关联的技术特点（按排序） """
    notes = ProjectImportantNote.query.filter_by(project_id=project_id).order_by(ProjectImportantNote.sort_order).all()
    result = [{"note_id": n.note.id, "label": n.note.label, "sort_order": n.sort_order} for n in notes]
    return ResponseTemplate.success(data=result, message="获取成功！")


@jwt_required()
def save_project_important_notes(project_id, data):
    """ 保存项目的技术特点（覆盖式更新） """
    try:
        # 先删除该项目已有的技术特点
        ProjectImportantNote.query.filter_by(project_id=project_id).delete()

        # 插入新数据
        for note_data in data:
            new_note = ProjectImportantNote(
                project_id=project_id,
                note_id=note_data["note_id"],
                sort_order=note_data["sort_order"]
            )
            db.session.add(new_note)

        db.session.commit()
        return ResponseTemplate.success(data=None, message="保存成功！")

    except Exception as e:
        db.session.rollback()
        logger.error(f"保存项目技术特点失败: {str(e)}")
        return ResponseTemplate.error(message="保存失败，请重试！")
