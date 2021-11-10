from PyQt5.QtWidgets import QApplication

from dll_power import CANMarathon
import pandas as pandas

from main import ExampleApp

app = QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

excel_data_df = pandas.read_excel('C:\\Users\\timofey.inozemtsev\\PycharmProjects\\VMULogger\\table_for_params.xlsx')
params_list = excel_data_df.to_dict(orient='records')
