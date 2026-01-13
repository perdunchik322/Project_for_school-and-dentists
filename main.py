from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidgetItem
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QHeaderView
from openpyxl import load_workbook
from PyQt6 import uic
import sys
import os

class ExcelProcessor:
    def __init__(self, path):
        self.path_to_file = path
        self.file = load_workbook(f"{path}", data_only=True)
        self.ws = self.file["Начисления"]
        self.table_name = "Итог"
        self.total_sum = 0
        self.global_minus = 0
        self.data_for_total = dict()

    def get_data(self):
        needed_columns = ["G", "H", "J", "I", "Y"]
        # for every rows
        for row in range(2, self.ws.max_row + 1):
            row_data = []
            # row from table
            for col in needed_columns:
                cell_value = self.ws[f"{col}{row}"].value
                row_data.append(cell_value)
            # calculating of quantity
            if None in row_data:  # minus from services and another things
                self.global_minus += row_data[-1]
            else:
                art = str(row_data[0])
                # make dictionary for unic arts
                if art not in self.data_for_total:
                    self.data_for_total[art] = [0, 0, 0, 0, 0, True]

                self.data_for_total[art][1] = row_data[1]
                sale = int(row_data[2])
                quantity = int(row_data[3])
                if sale > 0:
                    self.data_for_total[art][5] = False
                    self.data_for_total[art][0] += quantity
                elif sale < 0:
                    self.data_for_total[art][0] -= quantity
                self.data_for_total[art][4] += row_data[-1]
                if self.data_for_total[art][0] == 0:
                    del self.data_for_total[art]

        for i in self.data_for_total.keys():
            if self.data_for_total[i][0] > 0:
                self.data_for_total[i][3] = self.data_for_total[i][4]
            if self.data_for_total[i][0] <= 0:
                self.global_minus += self.data_for_total[i][4]
            del self.data_for_total[i][4]
            del self.data_for_total[i][4]
            self.total_sum += self.data_for_total[i][3]
            self.data_for_total[i][-1] = self.data_for_total[i][-1] + self.data_for_total[i][
                -1] * self.global_minus / self.total_sum
            if self.data_for_total[i][0] > 0:
                self.data_for_total[i][-2] = self.data_for_total[i][-1] / self.data_for_total[i][0]
        return self.data_for_total

    def write_data(self):
        # Удаляем лист, если он существует
        if self.table_name in self.file.sheetnames:
            std = self.file[self.table_name]
            self.file.remove(std)

        # Создаем новый лист
        ws = self.file.create_sheet(title=self.table_name)

        # Добавляем заголовки
        ws.append(["Артикул", "Название", "Кол-во", "Цена, шт", "Итого, So"])

        # Записываем данные
        for art in self.data_for_total:
            row = [
                art,
                self.data_for_total[art][1],
                self.data_for_total[art][0],
                self.data_for_total[art][2],
                self.data_for_total[art][3]
            ]
            ws.append(row)

        self.file.save(self.path_to_file)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui.ui', self)
        self.path_to_file = ''
        self.center_window()
        self.setup()

    def setup(self):
        self.choose_file_action.setShortcut("Ctrl+C")
        self.save_action.setShortcut("Ctrl+S")
        self.choose_file_action.triggered.connect(self.choose_file)
        self.save_action.triggered.connect(self.write_in_file)

    def choose_file(self):
        try:
            fname = QFileDialog.getOpenFileName(self, "Выбрать таблицу", '',
                                                            'Таблица (*.xlsx);;Таблица (*.xlsm);;Таблица (*.xls);;Все файлы(*)')
            self.path_to_file = fname[0]
            self.processor = ExcelProcessor(self.path_to_file)
            self.data_for_table = self.processor.get_data()
            self.init_view_table()
        except Exception:
            QMessageBox.critical(self, "Ошибка", "Ошибка обработки файла")

    def write_in_file(self):
        try:
            self.processor.write_data()
            self.statusBar().showMessage(f'Файл сохранён по пути:{self.path_to_file}')
            self.init_view_table()
        except Exception:
            QMessageBox.critical(self, "Ошибка", "Ошибка обработки файла")

    def init_view_table(self):
        self.view_table.setColumnCount(5)
        self.view_table.setHorizontalHeaderLabels(['Артикул', 'Название', 'Кол-во', 'Цена, шт', 'Итого'])
        self.view_table.setRowCount(len(self.data_for_table))
        header = self.view_table.horizontalHeader()

        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        for line_ind, art in enumerate(self.data_for_table):
            data = self.data_for_table[art]
            self.view_table.setItem(line_ind, 0, QTableWidgetItem(str(art)))
            self.view_table.setItem(line_ind, 1, QTableWidgetItem(str(data[1])))
            self.view_table.setItem(line_ind, 2, QTableWidgetItem(str(data[0])))
            self.view_table.setItem(line_ind, 3, QTableWidgetItem(str(data[2])))
            self.view_table.setItem(line_ind, 4, QTableWidgetItem(str(data[3])))

    def center_window(self):
        # moving window to center
        screen_geometry = self.screen().availableGeometry()  # get geom of screen
        window_geometry = self.frameGeometry()  # get geom of window
        center_point = screen_geometry.center()  # make variable for center of screen
        window_geometry.moveCenter(center_point)  # move center of window to screen center
        self.move(window_geometry.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())