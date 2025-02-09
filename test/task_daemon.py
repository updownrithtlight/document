import os
import json
import time
import logging
import pythoncom
import win32com.client

# 配置日志
logging.basicConfig(
    filename="task_daemon.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

TASK_FOLDER = r"C:\noneSystem\bj\document\uploads\output"  # 定义任务文件夹路径


def update_toc(doc_path):
    """
    使用 Word.Application 更新目录
    """
    pythoncom.CoInitialize()
    try:
        logging.info(f"开始处理文档：{doc_path}")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_path)
        if doc.TablesOfContents.Count > 0:
            doc.TablesOfContents(1).Update()  # 更新目录
        doc.Save()
        doc.Close()
        word.Quit()
        logging.info(f"✅ 文档处理完成：{doc_path}")
    except Exception as e:
        logging.error(f"❌ 文档处理失败：{e}")
    finally:
        pythoncom.CoUninitialize()


def run_daemon():
    """
    守护进程主循环
    """
    while True:
        try:
            # 遍历任务文件夹中的任务文件
            for filename in os.listdir(TASK_FOLDER):
                if filename.endswith(".json"):
                    task_file = os.path.join(TASK_FOLDER, filename)
                    with open(task_file, "r") as f:
                        task = json.load(f)
                        doc_path = task.get("doc_path")

                        if os.path.exists(doc_path):
                            update_toc(doc_path)  # 调用更新目录逻辑
                        else:
                            logging.warning(f"❌ 文档路径不存在：{doc_path}")

                    os.remove(task_file)  # 任务完成后删除任务文件
            time.sleep(5)  # 每隔 5 秒检查一次任务
        except Exception as e:
            logging.error(f"守护脚本运行出错：{e}")


if __name__ == "__main__":
    if not os.path.exists(TASK_FOLDER):
        os.makedirs(TASK_FOLDER)
    run_daemon()
