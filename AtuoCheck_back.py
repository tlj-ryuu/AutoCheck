from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPalette, QPixmap, QBrush, QImage
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QGraphicsScene
import AutoCheck
from file_module import *
import sys


class MainWindowActions(AutoCheck.Ui_Form, QMainWindow):

    def __init__(self):
        super(AutoCheck.Ui_Form, self).__init__()
        ##创建界面
        self.setupUi(self)
        self.file_path = ''
        self.files = []
        self.obj = None  # 当前打开的对象
        self.app = None  # 当前打开的app
        self.last_filename = '' # full path of last opened file
        self.signal = 0


        ##关联函数
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_label)
        self.pushButton.clicked.connect(self.get_filepath)
        self.pushButton_2.clicked.connect(self.confirm)
        self.pushButton_3.clicked.connect(self.check)

    def update_label(self):
        self.label_5.setText('')

    def get_filepath(self):
        self.file_path = self.lineEdit.text()
        print(self.file_path)
        self.files = get_filename(self.file_path)
        if self.files == [[], [], []]:
            self.label_5.setText('Please set correct path!')
        else:
            print(self.files)
            self.label_5.setText('File path has been set !')
            self.timer.start(4000)



    def check(self):
        if self.file_path == '':
            self.label_5.setText('Set file path first !')
            self.timer.start(4000)
        else:
            print('checking')
            self.label_5.setText('------Checking------')
            self.timer.start(4000)
            self.signal = self.spinBox.value()
            self.obj, self.app, self.last_filename = open_file(self.signal, self.obj, self.app, self.last_filename, self.files)
            if self.obj == None and self.app == None and self.last_filename == '':
                if self.signal == 1:
                    self.label_5.setText('Word file does not exist !')
                if self.signal == 2:
                    self.label_5.setText('Excel file does not exist !')
                if self.signal == 3:
                    self.label_5.setText('PPT file does not exist !')

    def confirm(self):
        if self.signal == 0:
            self.label_5.setText('Push the check btn first !')
            self.timer.start(4000)
        else:
            print('confirming')
            self.label_5.setText('------Confirming------')
            self.timer.start(4000)
            if self.signal == 1:
                self.obj, self.app = act_word(self.obj, self.app)
            elif self.signal == 2:
                self.obj, self.app = act_excel(self.obj, self.app)
            elif self.signal == 3:
                self.obj, self.app = act_ppt(self.obj, self.app)
            else:
                self.label_5.setText('Error occurred while confirming!')






if __name__ == '__main__':
    # 这里是界面的入口，在这里需要定义QApplication对象，之后界面跳转时不用再重新定义，只需要调用show()函数即可
    app = QApplication(sys.argv)
    # 显示创建的界面
    window = MainWindowActions()
    window.setWindowTitle('AutoCheck')
    window.show()

    sys.exit(app.exec_())