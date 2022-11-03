import sys
from PyQt5 import uic  # тут нет ошибки, pycharm видимо реально bad для pyqt5
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QToolTip, QMessageBox, QLabel, QWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon


# Окно с выгрузкой результата
class WindowGetResult(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WindowGetResult")


# Окно с основной работой программы (конвертирование)
class WindowWork(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WindowWork")


# Окно с загрузкой и изменением исходных файлов
class WindowDownload(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('app_2.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.setWindowIcon(QIcon('logo.png'))
        self.title = "Конвертор"
        self.return_btn.setFont(font)
        self.return_btn.clicked.connect(self.window_main_work)

        self.main_window()

    @staticmethod
    def input_files_in_programm():
        print('files had been inputed')

    def main_window(self):
        self.setWindowTitle(self.title)
        self.show()

    def window_main_work(self):
        self.win_main = WindowChoice()
        self.win_main.show()
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
        labels = [self.label_up, self.label_down]
        for label in labels:
            label.setFont(font)
        buttons = [self.exel_exel_btn, self.exel_word_btn, self.exel_json_btn]
        for btn in buttons:
            btn.setFont(font)
            btn.clicked.connect(self.window_download)
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
        self.win_d = WindowDownload()
        self.win_d.show()
        self.hide()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WindowChoice()
    sys.exit(app.exec())
