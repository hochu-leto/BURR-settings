"""
Цель этой ветки - создать  программу для просмотра параметров из кву цикл + и ПСТЭД цикл +
--------------------------------------------критично------------------------
- проверить на ПСТЭДе цикл+
--------------------------------------------хотелки-------------------------
- ограничить число записываемых параметров не более 30 штук, если в файле больше - брать только первые 50,
 либо предлагать выбрать другой файл
 - сделать просмотр также инвертора МЭИ
 - сделать просмотр также КВУ ТТС
---------------------------------------------непонятки----------------------
- не подключается когда подключили марафон или закрыли рткон или канвайс - тут вообще
какая-то непонятная фигня, есть подозрения, что это особенность работы марафона - без перезагрузки он периодически
отваливается . Выход - переход на квайзер и кан-хакер
--------------------------------------------исправил--------------------------------
- задел под парсинг файла с настройками от рткона - этим занимается отдельная прога
"""
import ctypes
import datetime
import pathlib
import struct

from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, pyqtSlot, QRegExp
from PyQt5.QtGui import QRegExpValidator, QIcon
from dll_power import CANMarathon
from PyQt5 import uic
from PyQt5.QtWidgets import QTableWidgetItem, QApplication, QMessageBox, QFileDialog, QMainWindow
import pandas as pandas


class VMU:
    def __init__(self, lst: tuple):
        self.req_id = lst[0]
        self.ans_id = lst[1]
        self.param_list = lst[2]
        self.req_list = self.feel_req_list()

    def feel_req_list(self):
        r_list = []
        for par in self.param_list:
            address = par['address']
            MSB = ((address & 0xFF0000) >> 16)
            LSB = ((address & 0xFF00) >> 8)
            sub_index = address & 0xFF
            # работает только для кву Цикл+
            data = [0x40, LSB, MSB, sub_index, 0, 0, 0, 0]
            r_list.append(data)
        return r_list

    def check_connection(self):
        # работает только для кву цикл +
        param_list = [[0x40, 0x18, 0x10, 0x02, 0x00, 0x00, 0x00, 0x00]]
        # Проверяю, есть ли подключение к кву
        # Так-то один параметр запрашивается, можно без реквест мэни, но это работает, осталю
        check = marathon.can_request_many(self.req_id, self.ans_id, param_list)
        return check


def make_vmu_params_list():
    global vmu
    fname = QFileDialog.getOpenFileName(window, 'Файл с нужными параметрами КВУ', dir_path + '//Tables',
                                        "Excel tables (*.xlsx)")[0]
    if fname and ('.xls' in fname):
        excel_data = pandas.read_excel(fname)
    else:
        print('File no choose')
        return False
    true_tuple = ('name', 'address', 'scale', 'unit')
    par_list = excel_data.to_dict(orient='records')
    check_list = par_list[0].keys()
    #  проверяю чтоб все нужные поля были в наличии
    result_list = [item for item in true_tuple if item not in check_list]

    if not result_list:
        if len(str(par_list[0]['address'])) < 6:
            QMessageBox.critical(window, "Ошибка ", 'Поле "address" должно быть\n'
                                                    '2 байта индекса + 1 байт сабиндекса\n'
                                                    'Например "0x520B09"',
                                 QMessageBox.Ok)
        else:
            new_vmu = fill_vmu_list(fname)
            if new_vmu:
                vmu = VMU(new_vmu)
            show_empty_params_list(vmu.param_list, 'vmu_param_table')
    else:
        QMessageBox.critical(window, "Ошибка ", 'В выбранном файле не хватает полей' + '\n' + ''.join(result_list),
                             QMessageBox.Ok)


def fill_vmu_list(file_name):
    need_fields = {'name', 'address', 'unit'}
    file = pandas.ExcelFile(file_name)
    sheet_name = file.sheet_names[0]
    if 'x' not in sheet_name and '-' not in sheet_name:
        QMessageBox.critical(None, "Ошибка ", 'Лист с параметрами назван неверно\n должны быть ID запроса-ответа '
                                              'нужного блока', QMessageBox.Ok)
        return False
    id_list = sheet_name.split('x')
    req_id = int(id_list[1][:-2], 16)
    ans_id = int(id_list[2], 16)
    sheet = file.parse(sheet_name=sheet_name)
    headers = list(sheet.columns.values)
    if not set(need_fields).issubset(headers):  # если в заголовках есть все нужные поля
        QMessageBox.critical(None, "Ошибка ", 'На листе не хватает столбцов\n name, address, unit', QMessageBox.Ok)
        return False
    vmu_params_list = sheet.to_dict(orient='records')
    exit_list = []
    for par in vmu_params_list:
        if str(par['name']) != 'nan':
            if str(par['address']) != 'nan':
                if isinstance(par['address'], str):
                    if '0x' in par['address']:
                        par['address'] = par['address'].rsplit('x')[1]
                    par['address'] = int(par['address'], 16)
                if str(par['scale']) == 'nan' or par['scale'] == 0:
                    par['scale'] = 1
                if str(par['scaleB']) == 'nan':
                    par['scaleB'] = 0
                exit_list.append(par)
    return req_id, ans_id, exit_list


def const_req_vmu_params():
    if not window.vmu_req_thread.running:
        window.vmu_req_thread.running = True
        window.thread_to_record.start()
    else:
        window.vmu_req_thread.running = False
        window.thread_to_record.terminate()


def adding_to_csv_file(name_or_value: str):
    data = []
    data_string = []
    for par in vmu.param_list:
        data_string.append(par[name_or_value])
    dt = datetime.datetime.now()
    dt = dt.strftime("%H:%M:%S.%f")
    if name_or_value == 'name':
        dt = 'time'
    data_string.append(dt)
    data.append(data_string)
    df = pandas.DataFrame(data)
    df.to_csv(window.vmu_req_thread.recording_file_name,
              mode='a',
              header=False,
              index=False,
              encoding='windows-1251')


def start_btn_pressed():
    # если записи параметров ещё нет, включаю ее
    if not window.record_vmu_params:
        window.vmu_req_thread.recording_file_name = pathlib.Path(pathlib.Path.cwd(),
                                                                 'VMU records',
                                                                 'vmu_record_' +
                                                                 datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") +
                                                                 '.csv')
        window.constantly_req_vmu_params.setChecked(True)
        window.constantly_req_vmu_params.setEnabled(False)
        window.connect_vmu_btn.setEnabled(False)
        window.record_vmu_params = True
        window.start_record.setText('Стоп')
        adding_to_csv_file('name')
    #  если запись параметров ведётся, отключаю её и сохраняю файл
    else:
        window.record_vmu_params = False
        window.start_record.setText('Запись')
        window.constantly_req_vmu_params.setChecked(False)
        window.constantly_req_vmu_params.setEnabled(True)
        window.connect_vmu_btn.setEnabled(True)
        # Reading the csv file
        file_name = str(window.vmu_req_thread.recording_file_name)
        df_new = pandas.read_csv(file_name, encoding='windows-1251')
        file_name = file_name.replace('.csv', '_excel.xlsx', 1)
        # saving xlsx file
        GFG = pandas.ExcelWriter(file_name)
        df_new.to_excel(GFG, index=False)
        GFG.save()
        QMessageBox.information(window, "Успешный Успех", 'Файл с записью параметров КВУ\n' +
                                'ищи в папке "VMU records"',
                                QMessageBox.Ok)


def connect_vmu():
    check = vmu.check_connection()
    if isinstance(check, list):
        check = check[0]
    if isinstance(check, str):
        QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + check, QMessageBox.Ok)
        window.connect_vmu_btn.setText('Подключиться')
        window.constantly_req_vmu_params.setEnabled(False)
        window.start_record.setEnabled(False)
        return False

    # запрашиваю список полученных ответов
    ans_list = marathon.can_request_many(vmu.req_id, vmu.ans_id, vmu.req_list)
    fill_vmu_params_values(ans_list)
    # отображаю сообщения из списка
    window.show_new_vmu_params()
    # разблокирую все кнопки и чекбоксы
    window.connect_vmu_btn.setText('Обновить')
    window.constantly_req_vmu_params.setEnabled(True)
    window.start_record.setEnabled(True)
    return True


def fill_vmu_params_values(ans_list: list):
    i = 0
    for par in vmu.param_list:
        message = ans_list[i]
        if not isinstance(message, str):
            value = (message[7] << 24) + \
                    (message[6] << 16) + \
                    (message[5] << 8) + message[4]
            if par['type'] == 'UNSIGNED8':
                par['value'] = ctypes.c_uint8(value).value
            elif par['type'] == 'UNSIGNED16':
                par['value'] = ctypes.c_uint16(value).value
            elif par['type'] == 'UNSIGNED32':
                par['value'] = ctypes.c_uint32(value).value
            elif par['type'] == 'SIGNED8':
                par['value'] = ctypes.c_int8(value).value
            elif par['type'] == 'SIGNED16':
                par['value'] = ctypes.c_int16(value).value
            elif par['type'] == 'SIGNED32':
                par['value'] = ctypes.c_int32(value).value
            elif par['type'] == 'FLOAT':
                par['value'] = struct.unpack('<f', bytearray(message[-4:]))[0]

            par['value'] = (par['value'] / par['scale'] - par['scaleB'])
            par['value'] = float('{:.2f}'.format(par['value']))
        i += 1


def show_empty_params_list(list_of_params: list, table: str):
    show_table = getattr(window, table)
    show_table.setRowCount(0)
    show_table.setRowCount(len(list_of_params))
    row = 0

    for par in list_of_params:
        name_item = QTableWidgetItem(par['name'])
        name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, 0, name_item)
        if str(par['description']) != 'nan':
            description = str(par['description'])
        else:
            description = ''
        description_item = QTableWidgetItem(description)
        show_table.setItem(row, 1, description_item)

        if par['address']:
            if str(par['address']) != 'nan':
                adr = hex(round(par['address']))
            else:
                adr = ''
            adr_item = QTableWidgetItem(adr)
            adr_item.setFlags(adr_item.flags() & ~Qt.ItemIsEditable)
            show_table.setItem(row, 2, adr_item)

        if str(par['unit']) != 'nan':
            unit = str(par['unit'])
        else:
            unit = ''
        unit_item = QTableWidgetItem(unit)
        unit_item.setFlags(unit_item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, show_table.columnCount() - 1, unit_item)

        value_item = QTableWidgetItem('')
        value_item.setFlags(value_item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, window.value_col, value_item)

        row += 1
    show_table.resizeColumnsToContents()


#  поток для опроса и записи в файл параметров кву
class VMUSaveToFileThread(QObject):
    running = False
    new_vmu_params = pyqtSignal(list)
    recording_file_name = ''

    # метод, который будет выполнять алгоритм в другом потоке
    def run(self):
        while True:
            if window.record_vmu_params:
                adding_to_csv_file('value')
            #  Получаю новые параметры от КВУ
            ans_list = []
            answer = marathon.can_request_many(vmu.req_id, vmu.ans_id, vmu.req_list)
            # Если происходит разрыв связи в блоком во время чтения
            #  И прилетает строка ошибки, то надо запихнуть её в список
            if isinstance(answer, str):
                ans_list.append(answer)
            else:
                ans_list = answer.copy()
            #  И отправляю их в основной поток для обновления
            self.new_vmu_params.emit(ans_list)

            response_time = window.response_time_edit.text()
            if response_time:
                response_time = int(response_time)
                if not response_time:
                    response_time = 1000
                if response_time < 10:
                    response_time = 10
                elif response_time > 60000:
                    response_time = 60000
            else:
                response_time = 1000
            # Это, конечно, жесть
            QThread.msleep(response_time)


class ExampleApp(QMainWindow):
    value_col = 3
    record_vmu_params = False

    def __init__(self):
        super().__init__()
        # Это нужно для инициализации нашего дизайна
        uic.loadUi('CANAnalyzer_2.ui', self)
        self.setWindowIcon(QIcon('icon.png'))
        # self.setupUi(self)
        #  Создаю поток для опроса параметров кву
        self.thread_to_record = QThread()
        # создадим объект для выполнения кода в другом потоке
        self.vmu_req_thread = VMUSaveToFileThread()
        # перенесём объект в другой поток
        self.vmu_req_thread.moveToThread(self.thread_to_record)
        # после чего подключим все сигналы и слоты
        self.vmu_req_thread.new_vmu_params.connect(self.add_new_vmu_params)
        # подключим сигнал старта потока к методу run у объекта, который должен выполнять код в другом потоке
        self.thread_to_record.started.connect(self.vmu_req_thread.run)

    @pyqtSlot(list)
    def add_new_vmu_params(self, list_of_params: list):
        # если в списке строка - нахер такой список, похоже, нас отсоединили
        # но бывает, что параметр не прилетел в первый пункт списка, тогда нужно проверить,
        # что хотя бы два пункта списка - строки( или придумать более изощерённую проверку)
        if len(list_of_params) == 1:  # or (isinstance(list_of_params[0], str) and isinstance(list_of_params[1],
            # str)):
            window.connect_vmu_btn.setText('Подключиться')
            window.connect_vmu_btn.setEnabled(True)
            window.start_record.setText('Запись')
            window.start_record.setEnabled(False)
            window.constantly_req_vmu_params.setChecked(False)
            window.constantly_req_vmu_params.setEnabled(False)
            window.record_vmu_params = False
            window.thread_to_record.running = False
            window.thread_to_record.terminate()
            QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + list_of_params[0], QMessageBox.Ok)
        else:
            fill_vmu_params_values(list_of_params)
            self.show_new_vmu_params()

    def show_new_vmu_params(self):
        row = 0
        for par in vmu.param_list:
            value_Item = QTableWidgetItem(str(par['value']))
            value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
            self.vmu_param_table.setItem(row, window.value_col, value_Item)
            row += 1


app = QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp
dir_path = str(pathlib.Path.cwd())
vmu_param_file = 'table_for_params.xlsx'
vmu = VMU(fill_vmu_list(pathlib.Path(dir_path, 'Tables', vmu_param_file)))
marathon = CANMarathon()
show_empty_params_list(vmu.param_list, 'vmu_param_table')
# главные кнопки для КВУ
window.connect_vmu_btn.clicked.connect(connect_vmu)
window.start_record.clicked.connect(start_btn_pressed)
window.constantly_req_vmu_params.toggled.connect(const_req_vmu_params)
# в окошке с задержкой опроса могут быть только цифры
reg_ex_2 = QRegExp("[0-9]{1,5}")
window.response_time_edit.setValidator(QRegExpValidator(reg_ex_2))
window.response_time_edit.setText('1000')
window.select_file_vmu_params.clicked.connect(make_vmu_params_list)

window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
