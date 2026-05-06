#使用PyQt实现Python GUI
import sys
from PyQt5.QtWidgets import *
from python_excel_word_pdf import convert

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 设置窗口标题和大小
        self.setWindowTitle("WORD_EXCEL_PDF")
        self.setGeometry(100,100,700,500)  # x, y, width, height
        
        
        self.moban_path_label = QLabel("输出路径：",self)
        self.moban_path_label.setGeometry(100,100,60,30)
        self.moban_path_text = QLineEdit(self)
        self.moban_path_text.setGeometry(160,100,400,30)
        self.moban_path_text_button = QPushButton("...", self)
        self.moban_path_text_button.setGeometry(560, 100, 30, 30)
        self.moban_path_text_button.clicked.connect(self.moban_path_text_button_clicked)
       
        self.moban_file_label = QLabel("模板文件：",self)
        self.moban_file_label.setGeometry(100,150,60,30)
        self.moban_file_text = QLineEdit(self)
        self.moban_file_text.setGeometry(160,150,400,30)
        self.moban_file_text_button = QPushButton("...", self)
        self.moban_file_text_button.setGeometry(560, 150, 30, 30)
        self.moban_file_text_button.clicked.connect(self.moban_file_text_button_clicked)


        self.mingdan_file_label = QLabel("名单文件：",self)
        self.mingdan_file_label.setGeometry(100,200,60,30)
        self.mingdan_file_text = QLineEdit(self)
        self.mingdan_file_text.setGeometry(160,200,400,30)
        self.mingdan_file_text_button = QPushButton("...", self)
        self.mingdan_file_text_button.setGeometry(560, 200, 30, 30)
        self.mingdan_file_text_button.clicked.connect(self.mingdan_file_text_button_clicked)

        self.zhanwei_label = QLabel("占位符",self)
        self.zhanwei_label.setGeometry(100,250,60,30)
        self.zhanwei_label_text = QLineEdit(self)
        self.zhanwei_label_text.setGeometry(160,250,400,30)

        self.n = QLabel("第几列",self)
        self.n.setGeometry(100,300,60,30)
        self.n_text = QLineEdit(self)
        self.n_text.setGeometry(160,300,400,30)

        self.button = QPushButton("START", self)
        self.button.setGeometry(100, 350, 100, 30)
        self.button.clicked.connect(self.button_clicked)


    def moban_path_text_button_clicked(self):
        folderPath = QFileDialog.getExistingDirectory(self, "选择模版文件夹")
        if folderPath:
            self.moban_path_text.setText(folderPath)
    def moban_file_text_button_clicked(self):
        filePath,_ = QFileDialog.getOpenFileName(self, "选择模版文件")
        if filePath:
            self.moban_file_text.setText(filePath)
    def mingdan_file_text_button_clicked(self):
        filePath,_ = QFileDialog.getOpenFileName(self, "选择名单文件")
        if filePath:
            self.mingdan_file_text.setText(filePath)
    
    def button_clicked(self):
        moban_path = self.moban_path_text.text()
        moban_file = self.moban_file_text.text()
        mingdan_file = self.mingdan_file_text.text()
        zhanwei = self.zhanwei_label_text.text()
        n = int(self.n_text.text())
        
        self.button.setText('处理中....')
        convert(moban_path,moban_file,mingdan_file,zhanwei,n-1)
        self.button.setText('处理完成')
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())