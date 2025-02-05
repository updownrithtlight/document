from app import db, jwt_required, logger
from app.exceptions.exceptions import CustomAPIException
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
        raise CustomAPIException("保存失败，请重试！", 404)



def get_important_notes(project_id):
    """ 获取项目关联的技术特点（按排序） """
    notes = ProjectImportantNote.query.filter_by(project_id=project_id).order_by(ProjectImportantNote.sort_order).all()
    # 构造数据列表，每个条目包含模板可直接使用的字段
    note_list = []
    for n in notes:
        note_list.append({
            "note_id": n.note.id,
            "label": n.note.label,
            "sort_order": n.sort_order
        })

    # 将 feature_list 包装为可直接在模板中使用的上下文
    # 如 docxtpl 中可以 {{ features }} 循环，也可以用 features.label
    result = {
        "important_notes": note_list
    }

    return result


