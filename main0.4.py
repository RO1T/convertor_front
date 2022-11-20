from PyQt5 import uic
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QDialog, QStackedWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon, QPixmap
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt
from datetime import datetime
from datetime import timedelta
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


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


class TableModel(QtCore.QAbstractTableModel):

    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def load_data(self, data):
        self.beginResetModel()
        self._data=data
        self.endResetModel()
    def data(self, index, role):
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            if isinstance(value, datetime):
                # Render time to YYY-MM-DD.
                return value.strftime("%Y-%m-%d")
            if isinstance(value, int):
                # Render time to YYY-MM-DD.
                return str(value)

            if isinstance(value, float):
                # Render float to 2 dp
                return "%.2f" % value

            if isinstance(value, str):
                # Render strings with quotes
                return '%s' % value
            return str(value)

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Vertical:
                return str(self._data.index[section])


class WorkWindow(QDialog):
    def __init__(self):
        super(WorkWindow, self).__init__()
        uic.loadUi('dialog_3.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        # данные с екселя
        filename = 'start.xlsx'
        df_input = pd.read_excel(filename, 'исходный формат')
        df_result = pd.read_excel(filename, 'нужный формат')
        self.convertor = Convertor(filename)
        # кнопки
        self.original_clear.clicked.connect(self.clear_orig)
        self.result_clear.clicked.connect(self.clear_res)
        self.apply_btn.clicked.connect(self.apply_changes)
        self.change_file_btn.clicked.connect(self.change_file_func)
        self.go_to_download_btn.clicked.connect(self.go_to_download_file_func)
        # таблица исходная
        self.model_original = TableModel(df_input)
        self.table_before.setModel(self.model_original)
        self.table_before.horizontalHeader().sectionClicked.connect(self.click_handler_original)
        # таблица итоговая
        self.model_result = TableModel(df_result)
        self.table_after.setModel(self.model_result)
        self.table_after.horizontalHeader().sectionClicked.connect(self.click_handler_result)

    def clear_orig(self):
        self.original.setText('')

    def clear_res(self):
        self.result.setText('')

    def click_handler_original(self, e):
        column_text = self.model_original.headerData(e, Qt.Horizontal, Qt.DisplayRole)
        self.original.setText(self.original.text() + column_text + ', ')

    def click_handler_result(self, e):
        column_text = self.model_result.headerData(e, Qt.Horizontal, Qt.DisplayRole)
        self.result.setText(self.result.text() + column_text + ', ')


    def apply_changes(self):
        command = self.get_command()
        self.convertor.execute(command)
        self.model_result.load_data(self.convertor.result)
        self.clear_res()
        self.clear_orig()


    def get_command(self):
        return (self.command.currentText(), self.original.text()[:-2].split(', '), self.result.text()[:-2].split(', '))

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
        # dfsdfsdfsdfsdf
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


class Convertor():
    def __init__(self, path):
        self.original = pd.read_excel(path, 'исходный формат')
        self.result = pd.read_excel(path, 'нужный формат')
        self.corr_fields = self.result.columns
        self.between = pd.DataFrame()

    def execute(self, command_data):
        command = command_data[0]
        input = command_data[1]
        corr = command_data[2]
        self.corr = corr
        if command == 'RENAME':
            self.rename(command_data)
            self.fill_result()
            self.fix_date()
            return
        changer = self.get_func(command)
        columns_for_change = [self.original[x] for x in input]
        corr_columns = changer(columns_for_change)
        # заполняет только первый столб, исправить!!!!
        for i in range(len(corr_columns[0]), self.original.shape[0]):
            for column in corr_columns:
                column.append(' ')
        for i in range(len(corr)):
            self.between[corr[i]] = corr_columns[i]
        self.fill_result()
        self.fix_date()

    def rename(self, command_data):
        dict_rename = {command_data[1][0]: command_data[2][0]}
        self.original.rename(columns=dict_rename, inplace=True)

    def fill_result(self):
        result = pd.DataFrame()
        for field in self.corr_fields:
            if self.between.__contains__(field):
                result[field] = self.between[field]
            elif self.original.__contains__(field):
                result[field] = self.original[field]
            else:
                result[field] = [' ' for i in range(len(result))]
        self.result = result

    def fix_date(self):
        for key in self.result.columns:
            if type(self.result[key][0]) is pd.Timestamp:
                list = []
                for item in self.result[key]:
                    a = str(item).replace('00:00:00', '')
                    list.append(a if a != 'NaT' else ' ')
                self.result[key] = list

    def get_func(self, command):
        if command == "SPLIT":
            return self.split_column
        elif command == "ZIP":
            return self.zip_columns
        else:  # command == "RENAME":
            return self.empty_method

    # [["фамилия"],["имя"],["отчество"]] -> [["ФИО"]] #corr_columns[i]
    def split_column(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.split_date(columns_for_change)
        splitter = ' '
        values = columns_for_change[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            # проверка на длину
            splitted_row = value.split(splitter)
            for i in range(len(splitted_row)):
                result[i].append(splitted_row[i])
        return result

    def split_date(self, columns):
        values = columns[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            # проверка на длину
            date = [datetime.date(value), datetime.time(value)]
            # (value.year,value.month,value.day)
            for i in range(len(date)):
                result[i].append(date[i])
        return result

    def zip_date(self, columns_for_change):
        result = []
        for i in range(len(columns_for_change[0])):
            date = columns_for_change[0][i] if type(columns_for_change[0][i]) is pd.Timestamp else \
                columns_for_change[1][i]
            time = columns_for_change[0][i] if type(columns_for_change[0][i]) is not pd.Timestamp else \
                columns_for_change[1][i]
            if pd.isna(date) or pd.isna(time):
                result.append(' ')
                continue
            delta = timedelta(hours=time.hour, minutes=time.minute, seconds=time.second)
            result.append(date + delta)
        return [result]

    def zip_columns(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.zip_date(columns_for_change)
        result = []
        values = []
        cars = pd.concat(columns_for_change).dropna().sort_index().astype('str').to_list()
        for car in cars:
            values.append(car)
            if (car.replace('.', '').isdigit()):
                model = ' '.join(values)
                values.clear()
                result.append(model)
        result = [result]  # -> список [["",""]]
        # result = [[result[x]] for x in range(len(result))] #-> список списков [[""],[""]]
        return result

    def empty_method(self):
        pass

    def as_text(self, val):
        if val is None:
            return ""
        return str(val)

    def to_exel(self):
        writer = pd.ExcelWriter('result1.xlsx',
                                engine='openpyxl')
        self.result.to_excel(writer, 'нужный формат', index=False)

        wb = Workbook()
        ws = wb.active
        # добавляем строчки дфа в опенпайексель
        for r in dataframe_to_rows(self.result, index=False, header=True):
            ws.append(r)
        # серега что-то тут нашаманил
        for column in ws.columns:
            length = max(len(self.as_text(cell.value)) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = length + 2
        wb.save("result1.xlsx")


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
