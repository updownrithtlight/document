import win32com.client
word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(r"C:\noneSystem\bj\document\test\test.docx")  # 替换为测试路径
doc.Close()
word.Quit()
