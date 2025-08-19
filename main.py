import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtCore import QSettings
from GUI.GUI import Ui_Dialog
import traceback
from cut_grade import cutGrade


class MyApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
 
        # Настройки
        self.settings = QSettings("MyCompany", "MyAppName")

        # Инициализация путей к файлам
        self.file1 = ""
        self.file2 = ""

        # Подключение кнопок
        self.ui.pushButton_1.clicked.connect(self.select_file_1)
        self.ui.pushButton_2.clicked.connect(self.select_file_2)
        self.ui.pushButton.clicked.connect(self.run_script)

        # Загрузка предыдущих значений
        self.load_settings()

    def load_settings(self):
        """Загрузка значений из настроек"""
        self.file1 = self.settings.value("file1", "")
        self.file2 = self.settings.value("file2", "")
        subblock = self.settings.value("subblock", "")

        self.ui.lineEdit.setText(self.file1)
        self.ui.lineEdit_2.setText(self.file2)
        self.ui.lineEdit_3.setText(subblock)

    def select_file_1(self):
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выбери Excel файл 1", filter="Excel Files (*.xlsx *.xls);;All Files (*)")
        if fname:
            self.file1 = fname
            self.ui.lineEdit.setText(fname)
            self.settings.setValue("file1", fname)

    def select_file_2(self):
        fname, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выбери Excel файл 2", filter="Excel Files (*.xlsx *.xls);;All Files (*)")
        if fname:
            self.file2 = fname
            self.ui.lineEdit_2.setText(fname)
            self.settings.setValue("file2", fname)

    def run_script(self):
        if not self.file1 or not self.file2:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Сначала выбери оба файла!")
            return

        try:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)

            subblock = self.ui.lineEdit_3.text()
            self.settings.setValue("subblock", subblock)

            cutGrade(pathFileAssay=df1, pathFileStrings=df2, EXPLORATION_BLOCK=subblock)

            QtWidgets.QMessageBox.information(self, "Успешно", "Скрипт выполнен успешно.")
        except Exception as e:
            traceback_str = traceback.format_exc()
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Ошибка:\n{e}\n\n{traceback_str}")


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MyApp()
    window.show()
    app.exec_()
