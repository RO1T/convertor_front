from PyQt5.QtWidgets import QMessageBox
class ExcpetionHandler:
    def __init__(self):
        self.msg = QMessageBox()

    def warning_choice_msg(self, type, message):
        self.msg.setWindowTitle(type)
        self.msg.setText(message)
        self.msg.setIcon(QMessageBox.Warning)
        self.msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        return self.msg.exec_()

    def critical_msg(self, type, message):
        self.msg.setWindowTitle(type)
        self.msg.setText(message)
        self.msg.setIcon(QMessageBox.Critical)
        self.msg.exec_()
