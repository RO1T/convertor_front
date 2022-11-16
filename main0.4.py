from PyQt5 import uic
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QDialog, QStackedWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon, QPixmap


class DownloadWindow(QDialog):
    def __init__(self):
        super(DownloadWindow, self).__init__()
        uic.loadUi('dialog_4.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.back_btn.setStyleSheet('''
        background-image : url(off.png);
        background-color : transparent;
        ''')
        self.download_btn.clicked.connect(self.download_fun)
        self.back_btn.clicked.connect(self.back_fun)

    def download_fun(self):
        print('Downloading files...')

    def back_fun(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)


class WorkWindow(QDialog):
    def __init__(self):
        super(WorkWindow, self).__init__()
        uic.loadUi('dialog_3.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.apply_btn.clicked.connect(self.apply_changes)
        self.apply_btn.text
        self.change_file_btn.clicked.connect(self.change_file_func)
        self.go_to_download_btn.clicked.connect(self.go_to_download_file_func)

    def apply_changes(self):
        print('Applying changes...')

    def change_file_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)

    def go_to_download_file_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() + 1)


class InputWindow(QDialog):
    def __init__(self):
        super(InputWindow, self).__init__()
        uic.loadUi('dialog_2.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        buttons = [self.change_btn, self.input_btn]
        for btn in buttons:
            btn.setFont(font)
        font_underline = self.change_btn.font()
        font_underline.setUnderline(True)
        self.change_btn.setFont(font_underline)
        self.change_btn.setStyleSheet("background-color: white")

        self.input_btn.clicked.connect(self.input_func)
        self.change_btn.clicked.connect(self.change_func)

    def input_func(self):
        # input files
        print('Inputing files...')
        # after input...
        #dfsdfsdfsdfsdf
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def change_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)


class MainWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('dialog_1.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        labels = [self.label_up, self.label_down]
        for label in labels:
            label.setFont(font)
        buttons = [self.exel_exel_btn, self.exel_word_btn, self.exel_json_btn]
        for btn in buttons:
            btn.setFont(font)
        self.exel_exel_btn.clicked.connect(self.exel_exel_btn_fun)
        self.exel_word_btn.clicked.connect(self.exel_word_btn_fun)
        self.exel_json_btn.clicked.connect(self.exel_json_btn_fun)

    def exel_exel_btn_fun(self):
        # если вдруг инициализация окна заранее не сработает
        # input_w = InputWindow()
        # widgets.addWidget(input_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def exel_word_btn_fun(self):
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def exel_json_btn_fun(self):
        widgets.setCurrentIndex(widgets.currentIndex() + 1)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widgets = QStackedWidget()

    main_w = MainWindow()
    input_w = InputWindow()
    work_w = WorkWindow()
    download_w = DownloadWindow()

    widgets.addWidget(main_w)
    widgets.addWidget(input_w)
    widgets.addWidget(work_w)
    widgets.addWidget(download_w)

    widgets.setGeometry(main_w.geometry())
    widgets.setWindowTitle('Конвертор')
    widgets.setWindowIcon(QIcon('logo.png'))
    widgets.show()

    try:
        sys.exit(app.exec_())
    except:
        print('Leaving')
