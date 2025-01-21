import subprocess
import os

# LibreOffice 可执行文件路径（请修改为你的实际路径）
SOFFICE_PATH = r"C:\noneSystem\bj\LibreOfficePortable\App\libreoffice\program\soffice.exe"

# 原始 .xls 文件和转换后的 .xlsx 文件路径
xls_file_path = r"/uploads/template/物料档案维护20250117.xls"  # 你的 .xls 文件路径
output_dir = r"/uploads/template/converted"  # 转换后的文件存放目录

# 确保输出目录存在
os.makedirs(output_dir, exist_ok=True)

# 执行 LibreOffice 命令进行转换
try:
    result = subprocess.run(
        [SOFFICE_PATH, "--headless", "--convert-to", "xlsx", xls_file_path, "--outdir", output_dir],
        capture_output=True,
        text=True
    )

    # 获取转换后的 .xlsx 文件路径
    xlsx_file_path = os.path.join(output_dir, os.path.basename(xls_file_path).replace(".xls", ".xlsx"))

    # 检查是否成功生成 .xlsx 文件
    if os.path.exists(xlsx_file_path):
        print(f"✅ 转换成功: {xlsx_file_path}")
    else:
        print(f"❌ 转换失败: {result.stderr}")

except Exception as e:
    print(f"❌ 发生错误: {str(e)}")
