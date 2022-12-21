from PyQt5 import uic
import sys
import pandas as pd
import json
from PyQt5.QtWidgets import (QApplication,
                             QDialog,
                             QStackedWidget,
                             QFileDialog,
                             QMessageBox,
                             QDesktopWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon
from PyQt5.QtCore import QAbstractTableModel, Qt
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class DownloadWindow(QDialog):
    def __init__(self, convertor, name_chose):
        super(DownloadWindow, self).__init__()
        self.name_chose = name_chose
        uic.loadUi('dialog_4.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.conv = convertor
        self.back_btn.setStyleSheet('''
        background-image : url(off.png);
        background-color : transparent;
        ''')
        self.download_btn.clicked.connect(self.download_fun)
        self.back_btn.clicked.connect(self.back_fun)

    def call_error(self):
        self.not_found = QMessageBox()
        self.not_found.setWindowTitle('Ошибка')
        self.not_found.setText('Вы должны указать, куда загружать файл!')
        self.not_found.setIcon(QMessageBox.Warning)
        self.not_found.move(
            self.mapToGlobal(self.rect().center() - self.not_found.rect().center())
        )
        self.not_found.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        return self.not_found.exec_()

    def download_fun(self):
        if self.name_chose == 'exel':
            self.path = QFileDialog.getSaveFileName(self, f"Куда сохранить файл?", "",
                                                    "Excel (*.xlsx *.xls)")
            self.file_name = self.path[0].split('/')[-1]
            self.file_path_abs = self.path[0]
            if self.file_path_abs == '':
                button = self.call_error()
                if button != QMessageBox.No:
                    self.download_fun()
            else:
                self.conv.to_exel(self.path[0])
        elif self.name_chose == 'json':
            self.path = QFileDialog.getSaveFileName(self, f"Куда сохранить файл?", "",
                                                    "Json (*.json)")
            self.file_name = self.path[0].split('/')[-1]
            self.file_path_abs = self.path[0]
            if self.file_path_abs == '':
                button = self.call_error()
                if button != QMessageBox.No:
                    self.download_fun()
            else:
                self.conv.to_json(self.path[0])

    def back_fun(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)
        widgets.removeWidget(self)


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def load_data(self, data):
        self.beginResetModel()
        self._data = data
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


class Convertor:
    def __init__(self, path):
        self.original = pd.read_excel(path, 'исходный формат')
        self.original =self.original.fillna('')
        self.result = pd.read_excel(path, 'нужный формат')
        self.corr_fields = self.result.columns
        self.between = pd.DataFrame()

    def execute(self, command_data):
        command = command_data[0]
        input = command_data[1]
        corr = command_data[2]
        self.corr = corr

        changer = self.get_func(command)
        columns_for_change = [self.original[x] for x in input]
        corr_columns = changer(columns_for_change)

        for i in range(len(corr_columns[0]), self.original.shape[0]):
            for column in corr_columns:
                column.append(' ')

        for i in range(len(corr)):
            self.between[corr[i]] = corr_columns[i]
        self.fill_result()
        self.fix_date()

    def rename(self, columns_for_change):
        return columns_for_change

    def fill_result(self):
        result = pd.DataFrame()
        for field in self.corr_fields:
            if self.between.__contains__(field):
                result[field] = self.between[field]
            elif self.original.__contains__(field):
                result[field] = self.original[field]
            else:
                result[field] = [' ' for i in range(self.original.shape[0])]
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
        elif command == "RENAME":
            return self.rename
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
            if value=='':
                continue
            splitted_row = value.split(splitter)
            for i in range(length):
                result[i].append(splitted_row[i])
        return result

    def split_date(self, columns):
        values = columns[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            if value=='':
                continue
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
        values = []
        for i in range(len(columns_for_change[0])):
            value = []
            for j in range(len(columns_for_change)):
                value.append(columns_for_change[j][i])
            values.append(' '.join(value))
        return [values]

    def have_empty_columns(self):
        for field in self.corr_fields:
            if self.between.__contains__(field):
                continue
            elif self.original.__contains__(field):
                continue
            else:
                return True
        return False

    def empty_method(self):
        pass

    def as_text(self, val):
        if val is None:
            return ""
        return str(val)

    def to_exel(self, path):
        writer = pd.ExcelWriter(path,
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
        wb.save(path)

    def to_json(self, path):
        res = self.result.to_json(orient="index")
        parsed = json.loads(res)
        with open(path, 'w', encoding='utf-8') as outfile:
            json.dump(parsed, outfile, indent=4, ensure_ascii=False)

class WorkWindow(QDialog):
    def __init__(self, filename, name_chose):
        super(WorkWindow, self).__init__()
        self.download_w = None
        self.name_chose = name_chose
        uic.loadUi('dialog_3.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        # данные с екселя
        self.filename = filename
        self.convertor = Convertor(self.filename)
        df_input = self.convertor.original
        df_result = self.convertor.result
        # кнопки
        self.original_clear.clicked.connect(self.clear_orig)
        self.result_clear.clicked.connect(self.clear_res)
        self.apply_btn.clicked.connect(self.apply_changes)
        self.change_file_btn.clicked.connect(self.change_file_func)
        self.go_to_download_btn.clicked.connect(self.changer)
        self.msg = QMessageBox()
        # таблица исходная
        self.model_original = TableModel(df_input)
        self.table_before.setModel(self.model_original)
        self.table_before.horizontalHeader().sectionClicked.connect(self.click_handler_original)
        # таблица итоговая
        self.model_result = TableModel(df_result)
        self.table_after.setModel(self.model_result)
        self.table_after.horizontalHeader().sectionClicked.connect(self.click_handler_result)

    def not_implemented_alert(self, message):
        self.msg.setWindowTitle('Ошибка')
        self.msg.setText(message)
        self.msg.setIcon(QMessageBox.Critical)

        self.msg.move(
            self.mapToGlobal(self.rect().center() - self.msg.rect().center())
        )

        self.msg.exec_()

    def empty_column_warning_no_yes(self, message):
        self.warning = QMessageBox()
        self.warning.setWindowTitle('Предупреждение')
        self.warning.setText(message)
        self.warning.setIcon(QMessageBox.Warning)
        self.warning.move(
            self.mapToGlobal(self.rect().center() - self.warning.rect().center())
        )
        self.warning.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        button = self.warning.exec_()
        if button == QMessageBox.Yes:
            self.go_to_download_file_func()

    def change_filled_warning_no_yes(self, message):
        self.warning = QMessageBox()
        self.warning.setWindowTitle('Предупреждение')
        self.warning.setText(message)
        self.warning.setIcon(QMessageBox.Warning)
        self.warning.move(
            self.mapToGlobal(self.rect().center() - self.warning.rect().center())
        )
        self.warning.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        button = self.warning.exec_()
        if button == QMessageBox.Yes:
            self.apply()
        else:
            pass
            # No

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
        no_split = command[0] == 'SPLIT' and (len(command[1]) > 1 or len(command[2]) == 1)
        no_rename = command[0] == 'RENAME' and (len(command[1]) > 1 or len(command[2]) > 1)
        no_zip = command[0] == 'ZIP' and (len(command[1]) == 1 or len(command[2]) > 1)
        change_filled = False
        for column in command[2]:
            if column in self.convertor.between.columns:
                change_filled = True
        if command[0] == '':
            self.not_implemented_alert('Вы ничего не сделали!')
        elif no_rename:
            self.not_implemented_alert(
                'Для выполнения команды RENAME в каждой таблице выберите только по одному столбцу!')
        elif no_split:
            self.not_implemented_alert('Для выполнения команды SPLIT в  исходной таблице выберите только один стобец!')
        elif no_zip:
            self.not_implemented_alert('Для выполнения команды ZIP в  итоговой таблице выберите только один стобец!')
        elif change_filled:
            self.change_filled_warning_no_yes(
                'Вы пытаетесь изменить уже заполненную колонку, уверены, что хотите продолжить?')
        else:
            self.apply()

    def apply(self):
        try:
            command = self.get_command()
            self.convertor.execute(command)
            self.model_result.load_data(self.convertor.result)
        except:
            self.not_implemented_alert('Данную функцию нельзя выполнить')
        finally:
            self.clear_res()
            self.clear_orig()

    def get_command(self):
        return self.command.currentText(), self.original.text()[:-2].split(', '), self.result.text()[:-2].split(', ')

    def change_file_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)
        widgets.removeWidget(self)

    def changer(self):
        if self.convertor.have_empty_columns():
            self.empty_column_warning_no_yes('У вас остались незаполненные колонки, уверены, что хотите продолжить?')
        else:
            self.go_to_download_file_func()

    def go_to_download_file_func(self):
        self.download_w = DownloadWindow(self.convertor, self.name_chose)
        widgets.addWidget(self.download_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)


class InputWindow(QDialog):
    def __init__(self, name_chose):
        super(InputWindow, self).__init__()
        self.name_chose = name_chose
        self.not_found = None
        self.work_w = None
        self.file_path_abs = None
        self.file_name = None
        self.file_path = None
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
        self.file_path = QFileDialog.getOpenFileName(self, f"Выберите файл {self.name_chose}", "",
                                                     "Excel (*.xlsx *.xls)")
        self.file_name = self.file_path[0].split('/')[-1]
        self.file_path_abs = self.file_path[0]
        if self.file_path_abs == '':
            self.not_found = QMessageBox()
            self.not_found.setWindowTitle('Ошибка')
            self.not_found.setText('Вы должны выбрать файл!\nЕсли вы передумали, нажмите No.')
            self.not_found.setIcon(QMessageBox.Warning)
            self.not_found.move(
                self.mapToGlobal(self.rect().center() - self.not_found.rect().center())
            )
            self.not_found.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            button = self.not_found.exec_()
            if button == QMessageBox.No:
                self.change_func()
            else:
                self.input_func()
        else:
            try:
                self.work_w = WorkWindow(self.file_path_abs, self.name_chose)
                widgets.addWidget(self.work_w)
                widgets.setCurrentIndex(widgets.currentIndex() + 1)
            except ValueError:
                self.msg = QMessageBox()
                self.msg.setWindowTitle('Ошибка')
                self.msg.setText('Не правильный исходный файл!')
                self.msg.setIcon(QMessageBox.Critical)

                self.msg.move(
                    self.mapToGlobal(self.rect().center() - self.msg.rect().center())
                )
                self.msg.exec_()

    def change_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)
        widgets.removeWidget(self)


class MainWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.input_w = None
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

        self.msg = QMessageBox()

    def exel_exel_btn_fun(self):
        self.input_w = InputWindow('exel')
        widgets.addWidget(self.input_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def exel_word_btn_fun(self):
        self.not_implemented_alert()

    def exel_json_btn_fun(self):
        self.input_w = InputWindow('json')
        widgets.addWidget(self.input_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def not_implemented_alert(self):
        self.msg.setWindowTitle('Ошибка')
        self.msg.setText('В процессе разработки!')
        self.msg.setIcon(QMessageBox.Critical)

        self.msg.move(
            self.mapToGlobal(self.rect().center() - self.msg.rect().center())
        )

        self.msg.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widgets = QStackedWidget()

    main_w = MainWindow()
    widgets.addWidget(main_w)

    widgets.setGeometry(main_w.geometry())

    widgets.setWindowTitle('Конвертор')
    widgets.setWindowIcon(QIcon('logo.png'))
    widgets.setMaximumSize(1160, 591)
    qtRectangle = widgets.frameGeometry()
    centerPoint = QDesktopWidget().availableGeometry().center()
    qtRectangle.moveCenter(centerPoint)
    widgets.move(qtRectangle.topLeft())

    widgets.show()

    sys.exit(app.exec_())
