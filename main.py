"""
Проблемы программы:
(- вылетает после установки часто используемых параметром, при этом сам параметр успевает изменить) - не подтвердилось
- очень долго сохраняет параметры из рейки в файл
(- не работает кнопка обновить когда уже подключено) - перепроверить
- не подключается когда подключили марафон или закрыли рткон или канвайс
- нет индикации работы когда считывает параметры
- выбрасывает ошибку при установке задней или передней оси(но ось задаёт)
- в списке осей нет индикации заводской настройки
- не записывает значения параметров из файла настроек в устройство
- не записывает порядок байт
- после сохранения параметров БУРР их нет в листе

Избавляемся от:
- (когда нет подключения при переключении на другую вкладку в основных параметрах пока
опрашивает каждый параметр, выбрасывает ошибку на каждый ) -  исправил
- (глобальных путей файлов *.xlsx) - исправил

Добавляю:
 - (папки для сохранения записей кву и для настроек БУРР) - добавил
"""

import datetime
import pathlib
import sys
from pprint import pprint

from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, pyqtSlot, QRegExp
from PyQt5.QtGui import QColor, QRegExpValidator

# sys.path.insert(1, 'C:\\Users\\timofey.inozemtsev\\PycharmProjects\\dll_power')
from dll_power import CANMarathon
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QTableWidgetItem, QApplication, QMessageBox, QFileDialog

import CANAnalyzer_ui
import pandas as pandas

Front_Wheel = 0x4F5
Rear_Wheel = 0x4F6
current_wheel = Front_Wheel

often_used_params = {
    'zone_of_insensitivity': {'scale': 100,
                              'value': 0,
                              'address': 103,
                              'min': 1,
                              'max': 5,
                              'unit': '%'},
    'warning_temperature': {'scale': 1,
                            'value': 0,
                            'address': 104,
                            'min': 30,
                            'max': 80,
                            'unit': u'\N{DEGREE SIGN}'},
    'warning_current': {'scale': 100,
                        'value': 0,
                        'address': 105,
                        'min': 10,
                        'max': 60,
                        'unit': 'A'},
    'cut_off_current': {'scale': 100,
                        'value': 0,
                        'address': 403,
                        'min': 20,
                        'max': 80,
                        'unit': 'A'}
}
'''
    Вот список байта аварий, побитно начиная от самого младшего:
    Битовое поле формируется от самого младшего к старшему биту справа на лево.
    Например 
    Бит 0 соответствует значению 0000 0001 или 0х01  ModFlt:1;  // 0  авария модуля 
    Бит 1 это 0000 0010 или 0x02         SCFlt:1;  // 1  кз на выходе
    Бит 2 это 0000 0100 или 0х04        HellaFlt:1;  // 2  авария датчика положение/калибровки
    Бит 3 это 0000 1000 или 0х08        TempFlt:1;  // 3  перегрев силового радиатора
    Бит 4 это 0001 0000 или 0x10        OvVoltFlt:1;  // 4  перенапряжение Udc
    Бит 5 это 0010 0000 или 0х20        UnVoltFlt:1;   // 5  понижение напряжения Udc
    Бит 6 это 0100 0000 или 0х40        OverCurrFlt:1;// 6  длительная токовая перегрузка
    Бит 7 это 1000 0000 или 0х80        RevErrFlt:1;   // 7  неправильная полярность DC-мотора
'''

errors_list = {0x1: 'авария модуля',
               0x2: 'кз на выходе',
               0x4: 'авария датчика положение/калибровки',
               0x8: 'перегрев силового радиатора',
               0x10: 'перенапряжение Udc',
               0x20: 'понижение напряжения Udc',
               0x40: 'длительная токовая перегрузка',
               0x80: 'неправильная полярность DC-мотора',
               }

rb_param_list = {
    'current_wheel': {'scale': 'nan',
                      'value': 0,
                      'address': 35},
    'byte_order': {'scale': 'nan',
                   'value': 0,
                   'address': 109},
}
compare_param_dict = {}
rtcon_vmu = 0x1850460E
vmu_rtcon = 0x594


def make_vmu_params_list():
    fname = QFileDialog.getOpenFileName(window, 'Файл с нужными параметрами КВУ', dir_path,
                                        "Excel tables (*.xlsx)")[0]
    if fname and ('.xls' in fname):
        excel_data = pandas.read_excel(fname)
    else:
        print('File no choose')
        return False
    true_tuple = ('name', 'address', 'scale', 'scaleB', 'unit')
    par_list = excel_data.to_dict(orient='records')
    check_list = par_list[0].keys()
    result_list = [item for item in true_tuple if item not in check_list]

    if not result_list:
        if len(str(par_list[0]['address'])) > 6:
            vmu_params_list = fill_vmu_list(fname)
            window.vmu_param_table.itemChanged.connect(check_connection)
            show_empty_params_list(vmu_params_list, 'vmu_param_table')
        else:
            QMessageBox.critical(window, "Ошибка ", 'Поле "address" должно быть\n'
                                                    '2 байта индекса + 1 байт сабиндекса\n'
                                                    'Например "0x520B09"',
                                 QMessageBox.Ok)
    else:
        QMessageBox.critical(window, "Ошибка ", 'В выбранном файле не хватает полей' + '\n' + ''.join(result_list),
                             QMessageBox.Ok)


def fill_vmu_list(file_name):
    excel_data_df = pandas.read_excel(file_name)
    vmu_params_list = excel_data_df.to_dict(orient='records')
    exit_list = []
    for par in vmu_params_list:
        if str(par['name']) != 'nan':
            if str(par['address']) != 'nan':
                if isinstance(par['address'], str):
                    if '0x' in par['address']:
                        par['address'] = par['address'].rsplit('x')[1]
                    par['address'] = int(par['address'], 16)
                if str(par['scale']) == 'nan':
                    par['scale'] = 1
                if str(par['scaleB']) == 'nan':
                    par['scaleB'] = 0
                exit_list.append(par)
    return exit_list


def check_response_time(item):
    pass


def const_req_vmu_params():
    if not window.vmu_req_thread.running:
        window.vmu_req_thread.running = True
        window.thread_to_record.start()
    else:
        window.vmu_req_thread.running = False
        window.thread_to_record.terminate()


def start_btn_pressed():
    if not window.record_vmu_params:
        window.constantly_req_vmu_params.setChecked(True)
        window.constantly_req_vmu_params.setEnabled(False)
        window.connect_vmu_btn.setEnabled(False)

        window.record_vmu_params = True
        window.start_record.setText('Стоп')
        window.vmu_req_thread.recording_file_name = pathlib.Path(dir_path,
                                                                 'VMU records',
                                                                 'vmu_record_' +
                                                                 datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") +
                                                                 '.csv')
        i = 0
        column = []
        data = []
        data_string = []
        for par in vmu_params_list:
            column.append(i)
            data_string.append(par['name'])
            i += 1
        data.append(data_string)
        df = pandas.DataFrame(data, columns=column)
        df.to_csv(window.vmu_req_thread.recording_file_name,
                  mode='w',
                  header=False,
                  index=False,
                  encoding='windows-1251')
    else:
        window.record_vmu_params = False
        window.start_record.setText('Запись')
        window.constantly_req_vmu_params.setChecked(False)
        window.constantly_req_vmu_params.setEnabled(True)
        window.connect_vmu_btn.setEnabled(True)


def feel_req_list(p_list: list):
    req_list = []
    for par in p_list:
        address = par['address']
        MSB = ((address & 0xFF0000) >> 16)
        LSB = ((address & 0xFF00) >> 8)
        sub_index = address & 0xFF
        data = [0x40, LSB, MSB, sub_index, 0, 0, 0, 0]
        req_list.append(data)
    return req_list


def connect_vmu():
    param_list = [[0x40, 0x18, 0x10, 0x02, 0x00, 0x00, 0x00, 0x00]]
    # Проверяю, есть ли подключение к кву
    check = marathon.can_request_many(rtcon_vmu, vmu_rtcon, param_list)
    if isinstance(check, list):
        check = check[0]
    if isinstance(check, str):
        QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + check, QMessageBox.Ok)
        window.connect_vmu_btn.setText('Подключиться')
        window.constantly_req_vmu_params.setEnabled(False)
        window.start_record.setEnabled(False)
        return False

    # запрашиваю список полученных ответов
    ans_list = marathon.can_request_many(rtcon_vmu, vmu_rtcon, req_list)
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
    for par in vmu_params_list:
        message = ans_list[i]
        if not isinstance(message, str):
            value = (message[7] << 24) + \
                    (message[6] << 16) + \
                    (message[5] << 8) + message[4]
            par['value'] = (value / par['scale']) - par['scaleB']
            par['value'] = float('{:.2f}'.format(par['value']))
        i += 1
    print('Новые параметры КВУ записаны ')


def write_all_from_file_to_device():
    if get_param(42):  # проверка что есть связь с блоком
        window.params_table_2.itemChanged.disconnect()
        for i in range(window.params_table_2.rowCount()):
            param_from_device = window.params_table_2.item(i, window.value_col)
            param_from_file = window.params_table_2.item(i, window.value_col + 1)
            if param_from_device:
                if param_from_device.text() == param_from_file.text():
                    break
            value_Item = QTableWidgetItem(param_from_file.text())
            param_from_file.setBackground(QColor('white'))
            window.params_table_2.setItem(i, window.value_col, value_Item)
            window.params_table_2.setItem(i, window.value_col + 1, param_from_file)
            address = get_address(window.params_table_2.item(i, 0))
            set_param(address, int(param_from_file.text()))
        window.params_table_2.itemChanged.connect(window.save_item)


def show_compare_list(compare_param_dict: dict):
    window.params_table_2.itemChanged.disconnect()
    for i in range(window.params_table_2.rowCount()):
        param_name = window.params_table_2.item(i, 0).text()
        param_from_device = window.params_table_2.item(i, window.value_col)
        if param_from_device:
            param_from_device = param_from_device.text()
        else:
            param_from_device = ''
        param_from_file = str(compare_param_dict[param_name])
        value_Item = QTableWidgetItem(param_from_file)
        value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
        if param_from_device != param_from_file:
            value_Item.setBackground(QColor('red'))

        window.params_table_2.setItem(i, window.value_col + 1, value_Item)
    window.params_table_2.itemChanged.connect(window.save_item)


def make_compare_list():
    global compare_param_dict
    fname = QFileDialog.getOpenFileName(window,
                                        'Файл с настройками БУРР-30',
                                        dir_path,
                                        "Excel tables (*.xlsx)")[0]
    if fname and ('.xls' in fname):
        excel_data = pandas.read_excel(fname)
    else:
        print('File no choose')
        return False
    par_list = excel_data.to_dict(orient='records')
    for param in par_list:
        if str(param['editable']) != 'nan':
            # value = str(param['value'].split(':')[0].replace(',', '.'))
            value = int(param['value'])
            compare_param_dict[param['name']] = value
    show_compare_list(compare_param_dict)
    window.load_to_device_button.setEnabled(True)


def check_connection():
    pass


def show_value(col_value: int, list_of_params: list, table: str):
    if get_param(42):  # проверка что есть связь с блоком
        global wr_err
        show_table = getattr(window, table)
        show_table.itemChanged.disconnect()

        row = 0

        for par in list_of_params:
            if (not par['value']) or (str(par['value']) == 'nan'):
                value = get_param(int(par['address']))
                par['value'] = value
            else:
                value = par['value']
            print(value)

            value_Item = QTableWidgetItem(str(value))

            if str(par['editable']) != 'nan':
                value_Item.setFlags(value_Item.flags() | Qt.ItemIsEditable)
                value_Item.setBackground(QColor('#D7FBFF'))
            else:
                value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)

            if str(par['strings']) != 'nan':
                value_Item.setStatusTip(str(par['strings']))
                value_Item.setToolTip(str(par['strings']))

            show_table.setItem(row, col_value, value_Item)

            row += 1
        show_table.resizeColumnsToContents()
        show_table.itemChanged.connect(window.save_item)


def show_empty_params_list(list_of_params: list, table: str):
    show_table = getattr(window, table)
    show_table.itemChanged.disconnect()
    show_table.setRowCount(0)
    show_table.setRowCount(len(list_of_params))
    row = 0

    for par in list_of_params:
        name_Item = QTableWidgetItem(par['name'])
        name_Item.setFlags(name_Item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, 0, name_Item)
        if str(par['description']) != 'nan':
            description = str(par['description'])
        else:
            description = ''
        description_Item = QTableWidgetItem(description)
        show_table.setItem(row, 1, description_Item)

        if par['address']:
            if str(par['address']) != 'nan':
                # if isinstance(par['address'], str):
                #     if '0x' in par['address']:
                #         par['address'] = par['address'].rsplit('x')[1]
                #     par['address'] = int(par['address'], 16)
                adr = hex(round(par['address']))
            else:
                adr = ''
            adr_Item = QTableWidgetItem(adr)
            adr_Item.setFlags(adr_Item.flags() & ~Qt.ItemIsEditable)
            show_table.setItem(row, 2, adr_Item)

        if str(par['unit']) != 'nan':
            unit = str(par['unit'])
        else:
            unit = ''
        unit_Item = QTableWidgetItem(unit)
        unit_Item.setFlags(unit_Item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, show_table.columnCount() - 1, unit_Item)

        row += 1
    show_table.resizeColumnsToContents()
    show_table.itemChanged.connect(window.save_item)


def update_param():
    if get_param(42):  # проверка что есть связь с блоком
        if window.tab_burr.currentWidget() == window.often_used_params:
            window.best_params()
        elif window.tab_burr.currentWidget() == window.editable_params:
            # я зачем-то раньше обновлял пустой список, сейчас это не нужно
            # show_empty_params_list(editable_params_list, 'params_table_2')
            show_value(window.value_col, editable_params_list, 'params_table_2')
            if compare_param_dict:
                show_compare_list(compare_param_dict)
        elif window.tab_burr.currentWidget() == window.all_params:
            param_list_clear()
            window.list_of_params_table(window.list_bookmark.currentItem())
        window.pushButton_2.setText('Обновить')
        font = QtGui.QFont()
        font.setBold(False)
        window.pushButton_2.setFont(font)
        window.groupBox_4.setEnabled(True)
        window.set_current_wheel.setEnabled(True)
        window.byte_order.setEnabled(True)
    else:
        window.pushButton_2.setText('Подключиться')
        font = QtGui.QFont()
        font.setBold(True)
        window.pushButton_2.setFont(font)
        window.groupBox_4.setEnabled(False)
        window.set_current_wheel.setEnabled(False)
        window.byte_order.setEnabled(False)


def param_list_clear():
    for param in params_list:
        param['value'] = 'nan'
    return True


def rb_clicked():
    global current_wheel
    if window.radioButton.isChecked():
        current_wheel = Front_Wheel
    elif window.radioButton_2.isChecked():
        current_wheel = Rear_Wheel
    update_param()
    return True


def get_address(name: str):
    for param in params_list:
        if str(name) == str(param['name']):
            return int(param['address'])
    return 'nan'


def check_param(address: int, value):  # если новое значение - часть списка, то
    int_type_list = ['UINT32', 'UINT16', 'INT32', 'INT16', 'DATE']
    for param in params_list:
        if str(param['address']) != 'nan':
            if param['address'] == address:  # нахожу нужный параметр
                if str(param['editable']) != 'nan':  # он должен быть изменяемым
                    if isinstance(value, int):  # и переменная - число
                        value = int(value)
                        # if int(param['max']) >= value >= int(param['min']):  # причём это число в зоне допустимого
                        return value  # ну тогда так у ж и быть - отдаём это число
                        # else:
                        #     print(f"param {value} is not in range from {param['min']} to {param['max']}")
                    else:
                        # отработка попадания значения из списка STR и UNION
                        print(f"value is not numeric {param['type']}")
                        # string_dict = {}
                        # for item in param['strings'].strip().split(';'):
                        #     if item:
                        #         it = item.split('-')
                        #         string_dict[it[1].strip()] = int(it[0].strip())
                        # return string_dict[value.strip()]
                        return 'nan'
                else:
                    print(f"can't change param {param['name']}")
    return 'nan'


def set_param(address: int, value: int):
    global wr_err
    wr_err = marathon.check_connection()
    if wr_err:
        QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + wr_err, QMessageBox.Ok)
        return False

    address = int(address)
    value = int(value)
    data = [value & 0xFF,
            ((value & 0xFF00) >> 8),
            0, 0,
            address & 0xFF,
            ((address & 0xFF00) >> 8),
            0x2B, 0x10]
    print(' Trying to set param in address ' + str(address) + ' to new value ' + str(value))
    effort = marathon.can_write(current_wheel, data)
    if not effort:
        print(f'Successfully updated param in address {address} into device')
        for param in params_list:
            if param['address'] == address:
                param['value'] = value
                print(f'Successfully written new value {value} in {param["name"]}')
                return True
    else:
        wr_err = "can't write param into device"
        QMessageBox.critical(window, "Ошибка ", wr_err + '\n' + effort, QMessageBox.Ok)
    return False


def get_param(address):
    global wr_err
    data = 'OK'
    wr_err = marathon.check_connection()
    if wr_err:
        QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + wr_err, QMessageBox.Ok)
        return False
    request_iteration = 3
    address = int(address)
    LSB = address & 0xFF
    MSB = ((address & 0xFF00) >> 8)
    for i in range(request_iteration):  # на случай если не удалось с первого раза поймать параметр,
        # делаем ещё request_iteration запросов
        data = marathon.can_request(current_wheel, current_wheel + 2, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
        if not isinstance(data, str):
            return (data[1] << 8) + data[0]
    wr_err = "can't read answer"
    QMessageBox.critical(window, "Ошибка ", wr_err + '\n' + data, QMessageBox.Ok)
    return False


def get_all_params():
    if get_param(42):
        for param in params_list:
            if str(param['address']) != 'nan':
                if str(param['value']) == 'nan':
                    param['value'] = get_param(address=int(param['address']))
    if not wr_err:
        return True
    print(wr_err)
    return False


def save_all_params():
    if get_all_params():
        file_name = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = 'Burr-30_' + file_name + '.xlsx'
        full_file_name = pathlib.Path(dir_path, 'Burr settings', file_name)
        pandas.DataFrame(params_list).to_excel(full_file_name, index=False)
        print(' Save file success')
        QMessageBox.information(window, "Успешный Успех", 'Файл настроек подключенного блока\n' +
                                file_name + '\nищи в папке "Burr settings"',
                                QMessageBox.Ok)
        return True
    print('Fail save file')
    QMessageBox.critical(window, "Ошибка ", 'Файл с настройками БУРР сохранить не удалось',
                         QMessageBox.Ok)
    return False


#  поток для опроса и записи в файл параметров кву
class VMUSaveToFileThread(QObject):
    running = False
    new_vmu_params = pyqtSignal(list)
    recording_file_name = pathlib.Path(pathlib.Path.cwd(),
                                       'VMU records',
                                       'vmu_record_' +
                                       datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") +
                                       '.csv')

    # метод, который будет выполнять алгоритм в другом потоке
    def run(self):
        while True:
            if window.record_vmu_params:
                columns = []
                data = []
                data_string = []
                for par in vmu_params_list:
                    columns.append(par['name'])
                    data_string.append(par['value'])
                data.append(data_string)
                df = pandas.DataFrame(data, columns=columns)
                df.to_csv(self.recording_file_name,
                          mode='a',
                          index=False,
                          header=False,
                          encoding='windows-1251')
            #  Получаю новые параметры от КВУ
            ans_list = []
            answer = marathon.can_request_many(rtcon_vmu, vmu_rtcon, req_list)
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
            QThread.msleep(response_time)


class ExampleApp(QtWidgets.QMainWindow, CANAnalyzer_ui.Ui_MainWindow):
    name_col = 0
    desc_col = 1
    value_col = 3
    combo_col = 999
    unit_col = 3
    record_vmu_params = False

    def __init__(self):
        super().__init__()
        # Это нужно для инициализации нашего дизайна
        self.setupUi(self)
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
        if len(list_of_params) == 1 or (isinstance(list_of_params[0], str) and isinstance(list_of_params[1], str)):
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
        for par in vmu_params_list:
            value_Item = QTableWidgetItem(str(par['value']))
            value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
            self.vmu_param_table.setItem(row, window.value_col, value_Item)
            row += 1

    def setting_current_wheel(self, item):
        global current_wheel
        if current_wheel == Front_Wheel:
            current_wheel = Rear_Wheel
            if str(get_param(42)) == 'nan':  # запрашиваю версию ПО у задней оси,
                current_wheel = Front_Wheel
                set_param(35, 3)  # если ответа нет, то можно переименовывать в заднюю
                current_wheel = Rear_Wheel
                self.radioButton_2.setChecked(True)
            else:
                current_wheel = Front_Wheel
                QMessageBox.critical(window, "Ошибка ", "Уже есть один задний блок", QMessageBox.Ok)
                return False
        elif current_wheel == Rear_Wheel:
            current_wheel = Front_Wheel
            if str(get_param(42)) == 'nan':  # запрашиваю версию ПО у передней оси,
                current_wheel = Rear_Wheel
                set_param(35, 2)  # если ответа нет, то можно переименовывать в переднюю
                current_wheel = Front_Wheel
                self.radioButton_2.setChecked(True)
            else:
                current_wheel = Rear_Wheel
                QMessageBox.critical(window, "Ошибка ", "Уже есть один передний блок", QMessageBox.Ok)
                return False

    def set_byte_order(self, item):
        if item:
            rb_toggled = QApplication.instance().sender()
            if rb_toggled == window.rb_big_endian:
                print('выбор биг-эндиан')
                set_param(109, 0)  # охренеть как тупо
            elif rb_toggled == window.rb_little_endian:
                print('выбор литл-эндиан')
                set_param(109, 1)

    def moved_slider(self, item):
        slider = QApplication.instance().sender()
        param = slider.objectName()
        value = item / often_used_params[param]['scale']
        label = getattr(self, 'lab_' + param)
        label.setText(str(value) + often_used_params[param]['unit'])

    def set_slider(self, item):
        slider = QApplication.instance().sender()
        param = slider.objectName()
        item = slider.value()
        value = item / often_used_params[param]['scale']
        address = often_used_params[param]['address']
        # print(f'New {param} is {item}')
        label = getattr(self, 'lab_' + param)
        label.setText(str(value) + often_used_params[param]['unit'])
        # надо сделать цикл раза три запихнуть параметр и проверить, если не получилось - предупреждение
        for i in range(marathon.max_iteration):
            if set_param(address, item):
                check_value = get_param(address)
                if check_value == item:
                    print('Checked changed value - OK')
                    label.setStyleSheet('background-color: green')
                    return True
                print(check_value)
        QMessageBox.critical(window, "Ошибка ", 'Что-то пошло не по плану,\n данные не записались',
                             QMessageBox.Ok)
        label.setStyleSheet('background-color: red')
        return False

    def best_params(self):
        self.lb_soft_version.setText('Версия ПО БУРР ' + str(get_param(42)))

        errors = get_param(0)
        errors_str = ''
        if str(errors) != 'nan':
            for err_nom, err_str in errors_list.items():
                if errors & err_nom:
                    errors_str += err_str + '\n'
        else:
            errors_str = 'Нет ошибок '
        self.tb_errors.setText(errors_str)

        self.set_front_wheel_rb.toggled.disconnect()
        self.set_rear_wheel_rb.toggled.disconnect()
        if current_wheel == Front_Wheel:
            self.set_front_wheel_rb.setChecked(True)
        elif current_wheel == Rear_Wheel:
            self.set_rear_wheel_rb.setChecked(True)
        self.set_front_wheel_rb.toggled.connect(self.setting_current_wheel)
        self.set_rear_wheel_rb.toggled.connect(self.setting_current_wheel)

        self.rb_big_endian.toggled.disconnect()
        self.rb_little_endian.toggled.disconnect()
        byte_order = get_param(109)
        if byte_order == 0:
            self.rb_big_endian.setChecked(True)
        elif byte_order == 1:
            self.rb_little_endian.setChecked(True)
        self.rb_big_endian.toggled.connect(self.set_byte_order)
        self.rb_little_endian.toggled.connect(self.set_byte_order)

        for name, par in often_used_params.items():
            par['value'] = get_param(int(par['address']))
            if par['scale'] != 'nan':
                slider = getattr(self, name)
                label = getattr(self, 'lab_' + name)
                if par['value'] != 'nan':
                    param = par['value']
                else:
                    param = par['max'] * par['scale']
                print(f'Param {name} is {param}')
                slider.valueChanged.disconnect()
                slider.setValue(param)
                slider.valueChanged.connect(self.set_slider)
                param = param / par['scale']
                label.setText(str(param) + par['unit'])
                label.setStyleSheet('background-color: white')

    def list_of_params_table(self, item):
        item = bookmark_dict[item.text()]
        show_empty_params_list(item, 'params_table')
        show_value(self.value_col, item, 'params_table')

    def save_item(self, item):
        table_param = QApplication.instance().sender()
        new_value = item.text()
        if not new_value:
            print('Value is empty')
            return False

        name_param = table_param.item(item.row(), self.name_col).text()
        if item.column() == self.value_col:
            address_param = get_address(name_param)
            if str(address_param) != 'nan':
                value = check_param(address_param, new_value)
                if str(value) != 'nan':  # прошёл проверку
                    for i in range(3):
                        if set_param(address_param, value):
                            check_value = get_param(address_param)
                            if check_value == value:
                                print('Checked changed value - OK')
                                table_param.item(item.row(), self.value_col).setBackground(QColor('green'))
                                return True
                            else:
                                table_param.item(item.row(), self.value_col).setBackground(QColor('red'))
                        else:
                            print("Can't write param into device")
                    if address_param == 35:  # если произошла смена рейки, нужно поменять адреса
                        if self.radioButton.isChecked():
                            self.radioButton_2.setChecked()
                        else:
                            self.radioButton.setChecked()
                        rb_clicked()
                        return True
                    return False
                else:
                    print("Param isn't in available range")
            else:
                print("Can't find this value")
            return False
        elif item.column() == self.desc_col:
            for param in params_list:
                if name_param == str(param['name']):
                    param['description'] = new_value
                    return True
            print("Can't find this value")
            return False
        else:
            print("It's impossible!!!")
            return False


app = QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

dir_path = str(pathlib.Path.cwd())

vmu_param_file = 'table_for_params.xlsx'
vmu_params_list = fill_vmu_list(pathlib.Path(dir_path, 'Tables', vmu_param_file))
# заполняю дату с адресами параметров из списка, который задаётся в файле
req_list = feel_req_list(vmu_params_list)

burr_param_file = 'burr_params.xls'
excel_data_df = pandas.read_excel(pathlib.Path(dir_path, 'Tables', burr_param_file))
params_list = excel_data_df.to_dict(orient='records')
bookmark_dict = {}
bookmark_list = []
prev_name = ''
wr_err = ''
editable_params_list = []
for param in params_list:
    if str(param['editable']) != 'nan':
        editable_params_list.append(param)
    if param['code'].count('.') == 2:
        param['address'] = int(param['address'])
        bookmark_list.append(param)
    elif param['code'].count('.') == 1:
        bookmark_dict[prev_name] = bookmark_list  # это словарь где все параметры по группам
        bookmark_list = []
        if prev_name:
            window.list_bookmark.addItem(prev_name)
        prev_name = param['name']

marathon = CANMarathon()
window.radioButton.toggled.connect(rb_clicked)
window.radioButton_2.toggled.connect(rb_clicked)

window.params_table.itemChanged.connect(window.save_item)
window.params_table_2.itemChanged.connect(window.save_item)
window.vmu_param_table.itemChanged.connect(check_connection)

window.pushButton.clicked.connect(save_all_params)
window.pushButton_2.clicked.connect(update_param)

window.list_bookmark.setCurrentRow(0)
show_empty_params_list(bookmark_dict[window.list_bookmark.currentItem().text()], 'params_table')
show_empty_params_list(editable_params_list, 'params_table_2')
show_empty_params_list(vmu_params_list, 'vmu_param_table')
window.vmu_param_table.itemChanged.disconnect()
reg_ex_2 = QRegExp("[0-9]{1,5}")
window.response_time_edit.setValidator(QRegExpValidator(reg_ex_2))
window.response_time_edit.setText('1000')

window.list_bookmark.itemClicked.connect(window.list_of_params_table)
window.params_table.resizeColumnsToContents()

for name, par in often_used_params.items():
    slider = getattr(window, name)
    slider.setMinimum(par['min'] * par['scale'])
    slider.setMaximum(par['max'] * par['scale'])
    slider.setPageStep(par['scale'])
    slider.setTracking(False)
    slider.setValue(par['min'])
    slider.sliderMoved.connect(window.moved_slider)
    slider.valueChanged.connect(window.set_slider)

    label = getattr(window, 'lab_' + name)
    label.setText(str(par['min']) + par['unit'])

window.set_front_wheel_rb.toggled.connect(window.setting_current_wheel)
window.set_rear_wheel_rb.toggled.connect(window.setting_current_wheel)

window.rb_big_endian.toggled.connect(window.set_byte_order)
window.rb_little_endian.toggled.connect(window.set_byte_order)

window.load_file_button.clicked.connect(make_compare_list)
window.load_to_device_button.clicked.connect(write_all_from_file_to_device)
window.connect_vmu_btn.clicked.connect(connect_vmu)
window.start_record.clicked.connect(start_btn_pressed)
window.constantly_req_vmu_params.toggled.connect(const_req_vmu_params)
window.response_time_edit.textEdited.connect(check_response_time)
window.select_file_vmu_params.clicked.connect(make_vmu_params_list)

window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
