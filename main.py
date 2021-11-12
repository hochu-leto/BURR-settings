import datetime
import sys
from pprint import pprint

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

sys.path.insert(1, 'C:\\Users\\timofey.inozemtsev\\PycharmProjects\\dll_power')
# C:\Users\timofey.inozemtsev\PycharmProjects\VMULogger\dll_power
from dll_power import CANMarathon
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QTableWidgetItem, QComboBox, QApplication, QMessageBox, QAbstractSlider, QFileDialog

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

rb_param_list = {'current_wheel': {'scale': 'nan',
                                   'value': 0,
                                   'address': 35},
                 'byte_order': {'scale': 'nan',
                                'value': 0,
                                'address': 109},
                 }
compare_param_dict = {}
rtcon_vmu = 0x1850460E
vmu_rtcon = 0x594


def connect_vmu():
    param_list = [[0x40, 0x18, 0x10, 0x02, 0x00, 0x00, 0x00, 0x00]]
    check = marathon.can_request_many(rtcon_vmu, vmu_rtcon, param_list)
    if not isinstance(check, list):
        QMessageBox.critical(window, "Ошибка ", 'Нет подключения' + '\n' + check, QMessageBox.Ok)
        return False
    # pprint(check)
    req_list =[]
    for par in vmu_params_list:
        address = par['address']
        MSB = ((address & 0xFF0000) >> 16)
        LSB = ((address & 0xFF00) >> 8)
        sub_index = address & 0xFF
        data = [0x40, LSB, MSB, sub_index, 0, 0, 0, 0]
        req_list.append(data)
    ans_list = marathon.can_request_many(rtcon_vmu, vmu_rtcon, req_list)
    for message in ans_list:
        if isinstance(message, list):
            for byte in message:
                print(hex(byte), end=' ')
            print()
            value = (message[4] << 24) + \
                    (message[5] << 16) + \
                    (message[6] << 8) + \
                    message[7]
            print(value)
        else:
            print(message)

def write_all_from_file_to_device():
    for i in range(window.params_table_2.rowCount()):
        param_from_device = window.params_table_2.item(i, 2)
        param_from_file = window.params_table_2.item(i, 3)
        if param_from_device != param_from_file:
            window.params_table_2.setItem(i, 2, param_from_file)


def show_compare_list(compare_param_dict: dict):
    window.params_table_2.itemChanged.disconnect()
    for i in range(window.params_table_2.rowCount()):
        param_name = window.params_table_2.item(i, 0).text()
        param_from_device = window.params_table_2.item(i, 2)
        if param_from_device:
            param_from_device = param_from_device.text()
        else:
            param_from_device = ''
        param_from_file = str(compare_param_dict[param_name])
        value_Item = QTableWidgetItem(param_from_file)
        value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
        if param_from_device != param_from_file:
            value_Item.setBackground(QColor('red'))

        window.params_table_2.setItem(i, 3, value_Item)
        window.params_table_2.itemChanged.connect(window.save_item)


def make_compare_list():
    global compare_param_dict
    fname = QFileDialog.getOpenFileName(window, 'Файл с настройками БУРР-30', 'C:\\Users\\timofey.inozemtsev'
                                                                              '\\PycharmProjects\\CAN_Analyzer',
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
    global wr_err
    show_table = getattr(window, table)
    show_table.itemChanged.disconnect()

    row = 0

    for par in list_of_params:
        if str(par['value']) == 'nan':
            value = get_param(int(par['address']))
            par['value'] = value
        else:
            value = par['value']

        if wr_err:
            #  надо выкидывать предупреждение, что адаптер не обнаружен
            # QMessageBox.critical(window, "Ошибка ", "Кажется, адаптер не подключен", QMessageBox.Ok)
            print(wr_err)
            wr_err = ''
            break
        else:
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

        if str(par['address']) != 'nan':
            if isinstance(par['address'], str):
                if '0x' in par['address']:
                    par['address'] = par['address'].rsplit('x')[1]
                par['address'] = int(par['address'], 16)
            adr = hex(par['address'])
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
        window.current_wheel.setEnabled(True)
        window.byte_order.setEnabled(True)
    else:
        window.pushButton_2.setText('Подключиться')
        font = QtGui.QFont()
        font.setBold(True)
        window.pushButton_2.setFont(font)
        window.groupBox_4.setEnabled(False)
        window.current_wheel.setEnabled(False)
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
        print(f'Successfully updated param in address {address} into devise')
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
        pandas.DataFrame(params_list).to_excel(file_name, index=False)
        print(' Save file success')
        return True
    print('Fail save file')
    return False


class ExampleApp(QtWidgets.QMainWindow, CANAnalyzer_ui.Ui_MainWindow):
    name_col = 0
    desc_col = 1
    value_col = 3
    combo_col = 999
    unit_col = 3

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна

    def set_current_wheel(self, item):
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
        print(slider.SliderAction())
        param = slider.objectName()
        item = slider.value()
        value = item / often_used_params[param]['scale']
        address = often_used_params[param]['address']
        print(f'New {param} is {item}')
        label = getattr(self, 'lab_' + param)
        label.setText(str(value) + often_used_params[param]['unit'])
        # надо сделать цикл раза три запихнуть параметр и проверить, если не получилось - предупреждение
        if set_param(address, item):
            check_value = get_param(address)
            if check_value == item:
                print('Checked changed value - OK')
                label.setStyleSheet('background-color: green')
                return True
        label.setStyleSheet('background-color: red')

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
            self.front_wheel.setChecked(True)
        elif current_wheel == Rear_Wheel:
            self.rear_wheel.setChecked(True)
        self.set_front_wheel_rb.toggled.connect(self.set_current_wheel)
        self.set_rear_wheel_rb.toggled.connect(self.set_current_wheel)

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

vmu_params_list = []
excel_data_df = pandas.read_excel('C:\\Users\\timofey.inozemtsev\\PycharmProjects\\VMULogger\\table_for_params.xlsx')
vmu_params_list = excel_data_df.to_dict(orient='records')

excel_data_df = pandas.read_excel('C:\\Users\\timofey.inozemtsev\\PycharmProjects\\VMULogger'
                                  '\\burr_30_forw_v31_27072021.xls')
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

window.set_front_wheel_rb.toggled.connect(window.set_current_wheel)
window.set_rear_wheel_rb.toggled.connect(window.set_current_wheel)

window.rb_big_endian.toggled.connect(window.set_byte_order)
window.rb_little_endian.toggled.connect(window.set_byte_order)

window.load_file_button.clicked.connect(make_compare_list)
window.load_to_device_button.clicked.connect(write_all_from_file_to_device)
window.connect_vmu_btn.clicked.connect(connect_vmu)
window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
