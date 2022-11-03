import sys
from PyQt5 import uic  # тут нет ошибки, pycharm видимо реально bad для pyqt5
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QToolTip, QMessageBox, QLabel, QWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon


# Окно с выгрузкой результата
class WindowGetResult(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Window22222")


# Окно с основной работой программы (конвертирование)
class WindowWork(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Window22222")


# Окно с загрузкой и изменением исходных файлов
class WindowDownload(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('app_1.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.setWindowIcon(QIcon('logo.png'))
        self.title = "WindowDownload"
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
        self.exel_exel_btn.clicked.connect(self.window_download)
        self.main_window()

    def main_window(self):
        self.setWindowTitle(self.title)
        self.show()

    def window_download(self):
        self.w = WindowChoice()
        self.w.show()
        self.hide()


# Начальное окно с выбором конвертора
class WindowChoice(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('app_1.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.setWindowIcon(QIcon('logo.png'))
        self.title = "Конвертор"
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
        self.exel_exel_btn.clicked.connect(self.window_download)
        self.exel_word_btn.clicked.connect(self.exel_word_func)
        self.exel_json_btn.clicked.connect(self.exel_json_func)
        self.main_window()

    @staticmethod
    def exel_exel_func():
        self.window_download()

    @staticmethod
    def exel_word_func():
        self.window_download()

    @staticmethod
    def exel_json_func():
        self.window_download()

    def main_window(self):
        self.setWindowTitle(self.title)
        self.show()

    def window_download(self):
        self.w = WindowDownload()
        self.w.show()
        self.hide()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WindowChoice()
    sys.exit(app.exec())
