# vocabulary_coda
要扫描项目字段并生成 requirements.txt 文件，你可以使用 Python 中的 pip freeze 命令。这个命令会列出当前安装的所有包及其版本，并且可以通过重定向操作符 > 将输出保存到一个文件中。你可以按照以下步骤执行：

打开命令行界面（如终端或命令提示符）。
切换到你的项目目录。
运行 pip freeze > requirements.txt 命令。
exe命令
pyinstaller --onefile --add-data "app;app" --add-data "instance;instance" --add-data "uploads;uploads" --add-data ".env;." --add-data "config.py;."  run.py
