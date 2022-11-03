import sys
from PyQt5 import uic  # тут нет ошибки, pycharm видимо реально bad для pyqt5
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QToolTip, QMessageBox, QLabel, QWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('app_1.ui', self)
        self.setWindowTitle('Конвертор')
        self.setWindowIcon(QIcon('logo.png'))
        # gilroy нету в базе PyQt5
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.label_down = self.findChild(QLabel, 'label_down')
        self.label_up = self.findChild(QLabel, 'label_up')
        labels = [self.label_up, self.label_down]
        for label in labels:
            label.setFont(font)

        self.exel_exel_btn = self.findChild(QPushButton, 'exel_exel_btn')
        self.exel_word_btn = self.findChild(QPushButton, 'exel_word_btn')
        self.exel_json_btn = self.findChild(QPushButton, 'exel_json_btn')
        buttons = [self.exel_exel_btn, self.exel_word_btn, self.exel_json_btn]
        for btn in buttons:
            btn.setFont(font)

        # btn.clicked может подчеркиваться, но это нормально,
        # PyCharm видимо не видит что это btn , так как они инициализируются во время runtime
        # (20-22 lines of code)
        # ((PyQt5-stubs убирают подчеркивания))
        self.exel_exel_btn.clicked.connect(self.exel_exel_func)
        self.exel_word_btn.clicked.connect(self.exel_word_func)
        self.exel_json_btn.clicked.connect(self.exel_json_func)



    @staticmethod
    def exel_exel_func():
        print('exel_exel')

    @staticmethod
    def exel_word_func():
        print('exel_word')

    @staticmethod
    def exel_json_func():
        print('exel_json')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName('asd')
    myApp = MyApp()
    myApp.show()
    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')
