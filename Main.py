import datetime
import os
import sys
from pathlib import Path

import pandas as pd
from PyQt6 import QtWidgets, QtGui
from PyQt6.QtWidgets import QApplication, QDialog

# 确保这里的 QTMainView 是你刚才生成的 PyQt6 版本的 UI 文件
import QTMainView
from toolCore import get_files_by_extension, TransPdfToDocx, TransDocToDocx, get_info_from_cells, word_tables_to_list

# 设置路径逻辑
if getattr(sys, 'frozen', None):
    Base_path = Path(sys._MEIPASS)
else:
    Base_path = Path.cwd()


# PyQt6 插件路径通常会自动识别，如果需要指定，请确保路径正确
# os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(Base_path.joinpath("resource/plugins/platforms"))

# 注意：这段代码在原文件中位于类之外，可能是废弃代码，但保留在此以防万一
def open_file_external(self):
    fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", os.getcwd(),
                                                               "All Files(*);;Text Files(*.txt)")
    print(fileName)
    print(fileType)


class MyWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = QTMainView.Ui_Dialog()
        self.ui.setupUi(self)

        # 信号连接
        self.ui.open_file.clicked.connect(self.open_file)
        self.ui.ok.clicked.connect(self.ok)
        self.ui.pdf_to_word.clicked.connect(self.pdf2word)
        self.ui.doc_to_docx.clicked.connect(self.doc2docx)
        self.ui.clear_output.clicked.connect(self.clearOutput)

    def open_file(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹", os.getcwd())
        if directory:  # 增加判断，防止用户取消选择导致设置空字符串
            self.ui.file_text.setText(directory)
            self.init_all()

    def addOutPut(self, text, color='black'):
        text_browser = self.ui.output
        f = QtGui.QTextCharFormat()
        f.setForeground(QtGui.QColor(color))
        text_browser.setCurrentCharFormat(f)
        text_browser.append(text)

    def getInput_Key_dict(self):
        text_edit = self.ui.input_key
        content = text_edit.toPlainText()
        content = content.replace("\n", "").replace(" ", "").replace("\t", "").strip()
        r1 = content.split(",")
        for index, i in enumerate(r1):
            if "，" in str(i):
                del r1[index]
                r1 += i.split("，")
        return r1

    def get_file_text(self):
        text_edit = self.ui.file_text
        content = text_edit.text()
        return content

    def clearOutput(self):
        self.ui.output.clear()

    def init_all(self):
        try:
            self.WORD_DIR = self.get_file_text()
            if not self.WORD_DIR:  # 如果目录为空则不继续
                return

            self.OUT_DIR = self.WORD_DIR + "/output"
            self.addOutPut("[INFO]输入目录:" + self.WORD_DIR, color='blue')
            if not os.path.exists(self.OUT_DIR):
                os.makedirs(self.OUT_DIR)
                self.addOutPut("[INFO]创建目录:" + self.OUT_DIR, color='blue')
            else:
                self.addOutPut("[INFO]发现输出目录:" + self.OUT_DIR, color='blue')
            # 获取所有文档
            self.init_files(self.WORD_DIR)
            for i in self.word_file_list:
                self.addOutPut("[INFO]" + os.path.basename(i), color='green')
            for i in self.pdf_file_list:
                self.addOutPut("[INFO]" + os.path.basename(i), color='orange')
            self.addOutPut("[INFO]word总数量:" + str(len(self.word_file_list)))
            self.addOutPut("[INFO]pdf总数量:" + str(len(self.pdf_file_list)))
        except Exception as e:
            self.addOutPut("[ERROR]错误:" + str(e), color='red')

    def init_files(self, dir_path):
        self.word_file_list = get_files_by_extension(dir_path, ".docx")
        self.pdf_file_list = get_files_by_extension(dir_path, '.pdf')

    def ok(self):
        try:
            self.clearOutput()
            self.keys_dict = self.getInput_Key_dict()
            self.init_all()

            if not hasattr(self, 'word_file_list'):
                self.addOutPut("[ERROR]请先选择有效的目录", color='red')
                return

            # 解析字段
            keys_dict = self.getInput_Key_dict()
            self.addOutPut("[INFO]解析到提取字段:" + str(keys_dict), color='blue')
            # 处理部分
            res = []
            for i in self.word_file_list:
                file_name = os.path.basename(i)
                cells = []
                try:
                    cells = word_tables_to_list(i)
                except Exception as e:
                    self.addOutPut("[WARN]警告:" + str(e), color='orange')
                    self.addOutPut("[WARN]转换失败:" + str(file_name), color='orange')
                # 读取失败启用第二轮读取
                if ("".join(cells)) == "":
                    try:
                        TransDocToDocx(i, i)
                        cells = word_tables_to_list(i)
                    except Exception as e:
                        print(i)
                        print(e)
                # 继续判断
                if ("".join(cells)) == "":
                    self.addOutPut("[WARN]读取失败:" + str(file_name), color='orange')
                else:
                    # 处理逻辑
                    info = get_info_from_cells(cells, keys_dict)
                    self.addOutPut("[INFO]解析到数据:" + str(info), color='green')
                    res.append(info)

            if res:
                data = pd.DataFrame(res)
                current_time = datetime.datetime.now()
                time_string = current_time.strftime("%Y_%m_%d_%H_%M_%S")

                out_file = self.OUT_DIR + '/' + time_string + ".xlsx"
                data.to_excel(out_file, index=False)
                self.addOutPut("[INFO]汇总结果输出成功,保存目录:" + str(out_file), color='blue')
            else:
                self.addOutPut("[WARN]未解析到任何数据", color='orange')

        except Exception as e:
            print(e)
            self.addOutPut("[ERROR]错误:" + str(e), color='red')

    def pdf2word(self):
        try:
            self.clearOutput()
            self.init_all()
            if not hasattr(self, 'pdf_file_list'): return

            for i in self.pdf_file_list:
                file_name = os.path.basename(i)
                try:
                    # 注意：os.path.join 是更安全的路径拼接方式，但这里为了保持原逻辑不动
                    TransPdfToDocx(i, self.OUT_DIR + "\\" + file_name.split(".")[0] + ".docx")
                    self.addOutPut("[INFO]ptf2word转换成功:" + file_name, color='green')
                except Exception as e:
                    self.addOutPut("[WARN]警告:" + str(e), color='orange')
                    self.addOutPut("[WARN]转换失败:" + str(file_name), color='orange')
        except Exception as e:
            self.addOutPut("[ERROR]错误:" + str(e), color='red')

    def doc2docx(self):
        try:
            self.clearOutput()
            self.init_all()
            if not hasattr(self, 'word_file_list'): return

            # 格式转换
            for i in self.word_file_list:
                file_name = os.path.basename(i)
                try:
                    TransDocToDocx(i, i)
                    self.addOutPut("[INFO]doc2docx转换成功:" + file_name, color='green')
                except Exception as e:
                    self.addOutPut("[WARN]警告:" + str(e), color='orange')
                    self.addOutPut("[WARN]转换失败:" + str(file_name), color='orange')
        except Exception as e:
            self.addOutPut("[ERROR]错误:" + str(e), color='red')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = MyWindow()
    MainWindow.show()
    # PyQt6 中 exec_() 已弃用，使用 exec()
    sys.exit(app.exec())
