from flask import request
from app import db, jwt_required, logger, app
from app.models.result import ResponseTemplate
from app.models.models import MaterialInfo
import os
import pandas as pd
from werkzeug.utils import secure_filename


@jwt_required()
def get_material_info_list():
    # 获取分页参数
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('pageSize', 10))

    # 获取搜索参数
    material_code = request.args.get('material_code', '').strip()
    material_name = request.args.get('material_name', '').strip()
    model_specification = request.args.get('model_specification', '').strip()

    # 构建查询
    query = MaterialInfo.query

    # 动态添加过滤条件
    if material_code:
        query = query.filter(MaterialInfo.material_code.ilike(f"%{material_code}%"))
    if material_name:
        query = query.filter(MaterialInfo.material_name.ilike(f"%{material_name}%"))
    if model_specification:
        query = query.filter(MaterialInfo.model_specification.ilike(f"%{model_specification}%"))

    # 分页
    materials_paginated = query.paginate(page=page, per_page=page_size, error_out=False)
    materials = materials_paginated.items

    # 转换为字典列表
    material_list = [material.to_dict() for material in materials]

    # 返回结果
    return ResponseTemplate.success(
        data={
            "materials": material_list,
            "totalElements": materials_paginated.total,
            "page": page,
            "pageSize": page_size,
            "pages": materials_paginated.pages
        },
        message='success'
    )


@jwt_required()
def get_material_info(material_id):
    material = MaterialInfo.query.get(material_id)  # 根据 ID 获取材料信息
    if not material:
        return ResponseTemplate.error(message='MaterialInfo not found')
    return ResponseTemplate.success(data=material.to_dict(), message='success')


@jwt_required()
def create_material_info(data):
    new_material = MaterialInfo(
        material_code=data['material_code'],
        material_name=data['material_name'],
        model_specification=data.get('model_specification'),  # 规格可选
        unit=data.get('unit')  # 单位可选
    )
    db.session.add(new_material)
    db.session.commit()
    return ResponseTemplate.success(message='MaterialInfo created successfully')


@jwt_required()
def update_material_info(material_id, data):
    material = MaterialInfo.query.get(material_id)
    if not material:
        return ResponseTemplate.error(message='MaterialInfo not found')

    material.material_code = data['material_code']
    material.material_name = data['material_name']
    material.model_specification = data.get('model_specification')  # 更新规格
    material.unit = data.get('unit')  # 更新单位
    db.session.commit()

    return ResponseTemplate.success(message='MaterialInfo updated successfully')


@jwt_required()
def delete_material_info(material_id):
    material = MaterialInfo.query.get(material_id)
    if not material:
        return ResponseTemplate.error(message='MaterialInfo not found')
    db.session.delete(material)
    db.session.commit()
    return ResponseTemplate.success(message='MaterialInfo deleted successfully')



import os
import subprocess
import pandas as pd
from flask import request
from flask_jwt_extended import jwt_required
from werkzeug.utils import secure_filename
from sqlalchemy.exc import SQLAlchemyError
from zipfile import BadZipFile
from app import db, logger
from app.models.result import ResponseTemplate
from app.models.models import MaterialInfo

UPLOAD_FOLDER =app.config['OUTPUT_FOLDER']
ALLOWED_EXTENSIONS = {"xls", "xlsx", "csv"}
LIBREOFFICE_PATH = r"C:\noneSystem\bj\LibreOfficePortable\App\libreoffice\program\soffice.exe"  # LibreOffice 可执行路径


def allowed_file(filename):
    """ 判断文件是否是允许的格式 """
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS




def convert_xls_to_xlsx(xls_path):
    """ 使用 LibreOffice 将 .xls 转换为 .xlsx """
    output_dir = os.path.dirname(xls_path)  # 目标文件所在目录
    xlsx_path = xls_path.replace(".xls", ".xlsx")  # 目标转换后的文件路径

    try:
        # **执行 LibreOffice 命令**
        result = subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "xlsx", xls_path, "--outdir", output_dir],
            capture_output=True,
            text=True,
            cwd=os.path.dirname(LIBREOFFICE_PATH)  # **确保 LibreOffice 运行目录正确**
        )

        # **打印 LibreOffice 执行日志**
        print(f"✅ LibreOffice stdout: {result.stdout}")
        print(f"❌ LibreOffice stderr: {result.stderr}")

        # **检查是否成功生成 .xlsx 文件**
        converted_file_path = os.path.join(output_dir, os.path.basename(xlsx_path))
        if os.path.exists(converted_file_path):
            print(f"✅ 转换成功: {converted_file_path}")
            return converted_file_path
        else:
            print(f"❌ LibreOffice 没有生成 .xlsx 文件: {converted_file_path}")
            return None

    except Exception as e:
        print(f"❌ LibreOffice 转换失败: {str(e)}")
        return None


@jwt_required()
def import_materials():
    """ 处理 Excel / CSV 文件上传并存储到数据库 """

    # **检查是否有文件**
    if "file" not in request.files:
        return ResponseTemplate.error(message="没有文件上传")

    file = request.files["file"]
    if file.filename == "":
        return ResponseTemplate.error(message="文件名不能为空")

    if not allowed_file(file.filename):
        return ResponseTemplate.error(message="不支持的文件格式，仅支持 Excel 或 CSV")

    # **保存文件**
    filename = secure_filename(file.filename)
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    file.save(file_path)

    # **如果是 .xls，转换为 .xlsx**
    if filename.endswith(".xls"):
        converted_file_path = convert_xls_to_xlsx(file_path)
        if converted_file_path is None:
            os.remove(file_path)
            return ResponseTemplate.error(message=".xls 文件转换失败，请检查 LibreOffice 是否安装")
        file_path = converted_file_path  # 使用转换后的文件

    # **解析文件**
    try:
        if file_path.endswith(".csv"):
            try:
                df = pd.read_csv(file_path, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding="gbk")
        else:
            try:
                df = pd.read_excel(file_path, engine="openpyxl", header=1)  # 读取文件，指定表头
                df.columns = df.columns.str.strip()  # 去掉列名的多余空格

                # **清理空值**
                df = df.where(pd.notnull(df), None)  # 将 NaN 转换为 None
            except BadZipFile:
                return ResponseTemplate.error(message="Excel 文件损坏或格式错误，请重新上传")

        # **预处理列名**
        df.columns = df.columns.str.strip()

        # **检查是否包含必要列**
        column_mapping = {
            "物料代码": "material_code",
            "物料名称": "material_name",
            "型号规格": "model_specification",
            "计量单位": "unit"
        }
        if not set(column_mapping.keys()).issubset(df.columns):
            return ResponseTemplate.error(
                message=f"文件缺少必要列，必须包含: {list(column_mapping.keys())}"
            )

        # **重命名列**
        df = df.rename(columns=column_mapping)

        # **批量插入数据**
        new_materials = [
            MaterialInfo(
                material_code=row["material_code"],
                material_name=row["material_name"],
                model_specification=row.get("model_specification", None),
                unit=row.get("unit", None),
            )
            for _, row in df.iterrows()
        ]

        db.session.bulk_save_objects(new_materials)
        db.session.commit()

        return ResponseTemplate.success(message="数据导入成功", data={"imported_rows": len(new_materials)})

    except SQLAlchemyError as e:
        db.session.rollback()
        logger.error(f"数据库错误: {str(e)}")
        return ResponseTemplate.error(message="数据库操作失败，请检查数据格式或联系管理员")

    except Exception as e:
        logger.error(f"数据导入失败: {str(e)}")
        return ResponseTemplate.error(message=f"导入失败: {str(e)}")

    finally:
        os.remove(file_path)  # **清理临时文件**
