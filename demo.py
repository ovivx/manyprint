from PySide2.QtWidgets import *
import sys
import os
from mainprint import Pinrt


document_files = []

outpath = ""

class mainwindow(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    
    def initUI(self):
        self.setWindowTitle("manyprint")
        self.setFixedSize(960,600)

        # 选择文件打印
        self.btn_select_document = QPushButton("选择文件",self)
        self.btn_select_document.move(5,5)
        self.btn_select_document.clicked.connect(self.select_document)


        #选择输出文件
        self.listwiget = QListWidget(self)
        self.listwiget.resize(930,450)
        self.listwiget.move(5,30)

        self.btn_outpath = QPushButton("输入文件夹",self)
        self.btn_outpath.move(5,500)
        self.btn_outpath.clicked.connect(self.select_outpath) 

        self.labe_outpath = QLabel("",self)
        self.labe_outpath.move(100,500) 
        self.labe_outpath.resize(200,30)      

        # 开始打印文件
        self.btn_start_print = QPushButton("开始打印",self)
        self.btn_start_print.move(400,500)
        self.btn_start_print.clicked.connect(self.start_print)


        self.show()

    def select_document(self):

        path,_ = QFileDialog.getOpenFileNames(self,"选择文件",os.getcwd(),"*.xlsx *.xls *.doc *.docx")

        global document_files

        document_files=path

        self.listwiget.addItems(path)

    def select_outpath(self):

        path = QFileDialog.getExistingDirectory()
        
        global outpath

        outpath = path

        self.labe_outpath.setText(path)


    def start_print(self):

        for file in document_files:

            if file.endswith(".docx"):

                prints = Pinrt()
                prints.wordPrint(file,outpath)

            elif file.endswith(".xlsx"):

                QMessageBox.information(self,"警告","该功能暂未实现")
            
            else:

                QMessageBox.information(self,"警告","故事还在，敬请期待")

        QMessageBox.information(self,"消息","pdf转换完成")


        # workname ="Word.Application"

        # word = Dispatch(workname)

        # word.Visible = 0
        # word.DisplayAlerts =0

        # for file in document_files:

        #     doc = word.Documents.Open(file)


        #     filename = os.path.splitext(os.path.split(file)[-1])[0]+".pdf"

        #     path = os.path.join(os.getcwd(),filename)

        #     doc.SaveAs(path,17)

        
        # doc.Close()

        # word.Quit()



if __name__ == "__main__":

    app = QApplication(sys.argv)
    window = mainwindow()
    sys.exit(app.exec_())