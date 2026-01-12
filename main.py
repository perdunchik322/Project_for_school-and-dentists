from openpyxl import load_workbook
import time

start_time = time.time()


class ExcelProcessor:
    def __init__(self, name_of_file):
        self.file = load_workbook(f"{name_of_file}", data_only=True)
        self.ws = self.file["Начисления"]

        self.total_sum = 0
        self.global_minus = 0
        self.data_for_total = dict()

    def get_data(self):
        needed_columns = ["G", "H", "J", "I", "Y"]
        #for every rows
        for row in range(2, self.ws.max_row + 1):
            row_data = []
            #row from table
            for col in needed_columns:
                cell_value = self.ws[f"{col}{row}"].value
                row_data.append(cell_value)
            #calculating of quantity
            if None in row_data: # minus from services and another things
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
            self.data_for_total[i][-1] = self.data_for_total[i][-1] + self.data_for_total[i][-1]* self.global_minus / self.total_sum
            if self.data_for_total[i][0] > 0:
                self.data_for_total[i][-2] = self.data_for_total[i][-1] / self.data_for_total[i][0]

processor = ExcelProcessor("Отчет_по_товарам_за_период_2025_11_24_2025_11_30.xlsx")
processor.get_data()

end_time = time.time()
execution_time = end_time - start_time
print(f"Время выполнения: {execution_time:.4f} секунд")