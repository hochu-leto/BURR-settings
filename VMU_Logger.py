"""
--------------------------------------------критично------------------------------
- блокировать любые нажатия, пока опрашиваются параметры(- нет индикации работы когда считывает параметры)
    не получается это сделать, потому что пока опрашиваются параметры внутри функции изображение формы изменить
    невозможно. Что-то делается с формой только после выхода из функции - один выход -переходить на параллельные потоки

- если изменилось дескрипшн - сохранять описание параметра в файл с дескрипшн

- выбрасывает ошибку при установке задней или передней оси(- НЕ МОЖЕТ ПЕРЕКЛЮЧИТЬ РЕЙКУ С ЗАДНЕЙ НА ПЕРЕДНЮЮ)
(- в списке осей нет индикации заводской настройки)

- не записывает значения параметров из файла настроек в устройство
- после сохранения параметров БУРР их нет в листе param_list
- автоматическое добавление слайдеров из списка
- в слайдеры добавить коэффициенты регулятора
- расширить диапазон зоны нечувствительности до 0,05%
- расширить диапазон токовый до 110А
- сделать токовые слайдеры одинаковыми по максимуму и минимуму
- сохраняет файл с настройками рейки даже если величин нет
- сделать нормальную проверку записываемого параметра
--------------------------------------------хотелки-------------------------
- сделать поиск параметра по описанию и названию
- вместо блокировки любых нажатий использовать другой поток для опроса параметров
- задел под парсинг файла с настройками от рткона
-- индекс и саб_индекс вместо адресов кву
--- исправить проверку на адрес, и если есть индекс и саб_индекс, брать их за основу и запихивать в адрес
-- скале_валуе - если есть, и не ноль, то запихиваем в скале))) - если это когда-то понадобится
- ограничить число записываемых параметров не более 30 штук, если в файле больше - брать только первые 50,
 либо предлагать выбрать другой файл
---------------------------------------------непонятки----------------------
- НЕ ЗАПИСЫВАЕТ ПОРЯДОК БАЙТ - не подтвердилось
(- вылетает после установки часто используемых параметром, при этом сам параметр успевает
изменить) - не подтвердилось
- очень долго сохраняет параметры из рейки в файл (- не работает кнопка обновить когда
уже подключено) - перепроверить
- не подключается когда подключили марафон или закрыли рткон или канвайс - тут вообще
какая-то непонятная фигня, есть подозрения, что это особенность работы марафона - без перезагрузки он периодически
отваливается . Выход - переход на квайзер и кан-хакер
--------------------------------------------исправил--------------------------------
- не обновляет параметры визуально
- не закрывает канал после записи параметра
- (когда нет подключения при переключении на другую вкладку в основных параметрах пока
опрашивает каждый параметр, выбрасывает ошибку на каждый ) -  исправил
- (глобальных путей файлов *.xlsx) - исправил
- (папки для сохранения записей кву и для настроек БУРР) - добавил
- (для уменьшения вероятности вылета во время чтения с кву сделать проверку не двух первых записей ответного списка,
 а сделать еррор_каунтер, который увеличивается при каждой строке в списке вместо значения, и
 если он превышать треть от длины всего списка, - значит нас отключили) - реализовал в dll_power
- (добавить время в файл параметров кву) - добавил
Нежелательно копировать область с настройками по внутренним датчикам,
 а также номером устройства, номерами плат и датой выпуска.
  Данный сектор находится в диапазоне адресов с 300 по 349 DEC.
 - СДЕЛАЛ
"""
import ctypes
import datetime
import pathlib

from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, pyqtSlot, QRegExp
from PyQt5.QtGui import QColor, QRegExpValidator

# sys.path.insert(1, 'C:\\Users\\timofey.inozemtsev\\PycharmProjects\\dll_power')
from dll_power import CANMarathon
from PyQt5 import QtWidgets, QtGui, uic
from PyQt5.QtWidgets import QTableWidgetItem, QApplication, QMessageBox, QFileDialog
import CANAnalyzer_ui
import pandas as pandas


class Wheel():
    req_id = 0
    ans_id = 0


Front_Wheel = 0x4F5
Rear_Wheel = 0x4F6
current_wheel = Front_Wheel

often_used_params = {
    'zone_of_insensitivity': {'scale': 100,
                              'value': 0,
                              'address': 103,
                              'min': 0,
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
                        'min': 20,
                        'max': 100,
                        'unit': 'A'},
    'cut_off_current': {'scale': 100,
                        'value': 0,
                        'address': 403,
                        'min': 20,
                        'max': 100,
                        'unit': 'A'}
}
'''
    Вот список байта аварий, побитно начиная от самого младшего:
    Битовое поле формируется от самого младшего к старшему биту справа на лево.
    Например 
    Бит 0 это 0000 0001 или 0х01        ModFlt:1;       // 0  авария модуля 
    Бит 1 это 0000 0010 или 0x02        SCFlt:1;        // 1  кз на выходе
    Бит 2 это 0000 0100 или 0х04        HellaFlt:1;     // 2  авария датчика положение/калибровки
    Бит 3 это 0000 1000 или 0х08        TempFlt:1;      // 3  перегрев силового радиатора
    Бит 4 это 0001 0000 или 0x10        OvVoltFlt:1;    // 4  перенапряжение Udc
    Бит 5 это 0010 0000 или 0х20        UnVoltFlt:1;    // 5  понижение напряжения Udc
    Бит 6 это 0100 0000 или 0х40        OverCurrFlt:1;  // 6  длительная токовая перегрузка
    Бит 7 это 1000 0000 или 0х80        RevErrFlt:1;    // 7  неправильная полярность DC-мотора
небольшое дополнение - список ошибок был неполон. Вот полный
0- ModFlt ;             // 0  авария модуля
1- SC1Flt ;             // 1  кз на выходе
2- HellaFlt ;           // 2  авария датчика положение/калибровки
3- DINstopFlt ;         // 3  DINstopFlt 
4- REZERV 2 ;           // 4  REZERV 2
5- NoMoveFlt ;          // 5  NoMoveFlt   
6- I2tFlt ;             // 6  I2tFlt
7- TempFlt ;            // 7  перегрев силового радиатора
8- OpenMotor ;          // 8  OpenMotor
9- OvVoltFlt ;          // 9  перенапряжение Udc
10- UnVoltFlt ;         // 10 понижение напряжения Udc
11- NoCAN ;             // 11 Нет CANa
12- NoMB ;              // 12 Нет МодБаса
13- OverCurrFlt ;       // 13 длительная токовая перегрузка
14- RevErrFlt ;         // 14 неправильная полярность DC-мотора
15- REZERV_15 ;         // 15 REZERV_15 

'''

errors_list = {0x1: 'авария модуля',
               0x2: 'кз на выходе',
               0x4: 'авария датчика положени/калибровки',
               0x8: 'DIN stop Fault',
               0x10: 'REZERV',
               0x20: 'No Move Fault',
               0x40: 'I2t Fault',
               0x80: 'перегрев силового радиатора',
               0x100: 'Open Motor Fault',
               0x200: 'перенапряжение Udc',
               0x400: 'понижение напряжения Udc',
               0x800: 'нет CANa',
               0x1000: 'нет МодБаса',
               0x2000: 'длительная токовая перегрузка',
               0x4000: 'неправильная полярность DC-мотора',
               0x8000: 'REZERV 15',

               }

compare_param_dict = {}
rtcon_vmu = 0x1850460E
vmu_rtcon = 0x594


def change_current_wheel(target_wheel: int):
    global current_wheel
    # защита от повторного входа в функцию
    if target_wheel == has_wheel(current_wheel):
        return True
    # надо проверить на всех рейках может быть любое значение 0,1,2,3,4 , апрол
    err = ''
    if target_wheel == 2:
        if has_wheel(Front_Wheel):
            err = "Уже есть один передний блок"
    elif target_wheel == 3:
        if has_wheel(Rear_Wheel):
            err = "Уже есть один задний блок"
    else:
        err = 'Рейка должна быть или передней = 2, или задней = 3'

    if err:
        QMessageBox.critical(window, "Ошибка ", err, QMessageBox.Ok)
        return False

    err = marathon.can_write(current_wheel, [target_wheel, 0, 0, 0, 0x23, 0, 0x2B, 0x10])
    if not err:
        window.radioButton.toggled.disconnect()
        window.radioButton_2.toggled.disconnect()
        if current_wheel == Front_Wheel:
            current_wheel = Rear_Wheel
            window.radioButton_2.setChecked(True)
        else:
            current_wheel = Front_Wheel
            window.radioButton.setChecked(True)
        window.radioButton.toggled.connect(rb_clicked)
        window.radioButton_2.toggled.connect(rb_clicked)

        if target_wheel == has_wheel(current_wheel):
            for param in params_list:
                if param['address'] == 35:
                    param['value'] = target_wheel
                    return True
            err = 'Не найден параметр в списке - ерунда'
        else:
            err = f'Текущий параметр из устройства отличается от желаемого {target_wheel}'
    QMessageBox.critical(window, "Ошибка ", err, QMessageBox.Ok)
    return False


def show_waiting_tab():
    window.recent_tab = window.CAN.currentWidget()
    window.CAN.setEnabled(False)
    window.CAN.addTab(window.hidden_tab, '')
    window.CAN.setCurrentWidget(window.Wait_for_read)
    window.show()


def hide_waiting_tab():
    window.CAN.removeTab(window.CAN.currentIndex())
    window.CAN.setCurrentWidget(window.recent_tab)
    window.CAN.setEnabled(True)


def make_vmu_params_list():
    fname = QFileDialog.getOpenFileName(window, 'Файл с нужными параметрами КВУ', dir_path,
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
            vmu_params_list = fill_vmu_list(fname)
            window.vmu_param_table.itemChanged.connect(check_connection)
            show_empty_params_list(vmu_params_list, 'vmu_param_table')
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
    for par in vmu_params_list:
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
            #  вывод на печать полученных ответов
            # print(par['name'])
            # for j in message:
            #     print(hex(j), end=' ')
            # если множителя нет, то берём знаковое int16
            if par['scale'] == 1:
                par['value'] = ctypes.c_int16(value).value
            # возможно, здесь тоже нужно вытаскивать знаковое int, ага, int32
            else:
                value = ctypes.c_int32(value).value
                # print(' = ' + str(value), end=' ')
                par['value'] = (value / par['scale'])
                # print(' = ' + str(par['value']))
            par['value'] = float('{:.2f}'.format(par['value']))
        i += 1
    print('Новые параметры КВУ записаны ')


# очень опасная процедура - надо очень пристально ее проверить
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
    if update_connect_button():  # проверка что есть связь с блоком
        show_table = getattr(window, table)

        row = 0

        for par in list_of_params:
            if (not par['value']) or (str(par['value']) == 'nan'):
                value = get_param(int(par['address']))
                par['value'] = value
            else:
                value = par['value']

            value_Item = QTableWidgetItem(str(value))

            if str(par['editable']) != 'nan':
                value_Item.setFlags(value_Item.flags() | Qt.ItemIsEditable)
                value_Item.setBackground(QColor('#D7FBFF'))
            else:
                value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)

            if str(par['strings']) != 'nan':
                value_Item.setStatusTip(str(par['strings']))
                value_Item.setToolTip(str(par['strings']))
            show_table.itemChanged.disconnect()
            show_table.setItem(row, col_value, value_Item)
            show_table.itemChanged.connect(window.save_item)

            row += 1
        show_table.resizeColumnsToContents()
    marathon.close_marathon_canal()


def show_empty_params_list(list_of_params: list, table: str):
    show_table = getattr(window, table)
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

        value_Item = QTableWidgetItem('')
        value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
        show_table.setItem(row, window.value_col, value_Item)

        row += 1
    show_table.resizeColumnsToContents()


def update_param():
    if update_connect_button():  # проверка что есть связь с блоком
        if window.tab_burr.currentWidget() == window.often_used_params:
            window.best_params()
        elif window.tab_burr.currentWidget() == window.editable_params:
            param_list_clear()
            show_value(window.value_col, editable_params_list, 'params_table_2')
            if compare_param_dict:
                show_compare_list(compare_param_dict)
        elif window.tab_burr.currentWidget() == window.all_params:
            param_list_clear()
            show_value(window.value_col, bookmark_dict[window.list_bookmark.currentItem().text()], 'params_table')


def update_connect_button():
    software_version = get_param(42)
    print('software_version = ' + str(software_version))
    if software_version:
        window.pushButton_2.setText('Обновить')
        font = QtGui.QFont()
        font.setBold(False)
        window.pushButton_2.setFont(font)
        return software_version

    window.pushButton_2.setText('Подключиться')
    font = QtGui.QFont()
    font.setBold(True)
    window.pushButton_2.setFont(font)
    window.groupBox_4.setEnabled(False)
    window.set_current_wheel.setEnabled(False)
    window.byte_order.setEnabled(False)
    return False


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


# какая-то хрень а не проверка
def check_param(address: int, value):
    if isinstance(value, int):
        return int(value)
    return 'nan'


def set_param(address: int, value: int):
    # Здесь всё нахрен надо исправить
    address = int(address)
    value = int(value)
    if address == 0x23:
        change_current_wheel(value)
        return True
    LSB = address & 0xFF
    MSB = ((address & 0xFF00) >> 8)
    data = [value & 0xFF,
            ((value & 0xFF00) >> 8),
            0, 0, LSB, MSB,
            0x2B, 0x10]
    print(' Trying to set param in address ' + str(address) + ' to new value ' + str(value))
    err = marathon.can_write(current_wheel, data)
    if not err:
        data = marathon.can_request(current_wheel, current_wheel + 2, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
        if not isinstance(data, str):
            data = ((data[1] << 8) + data[0])
            if value == data:
                print(f'Successfully updated param in address {address} into device')
                for param in params_list:
                    if param['address'] == address:
                        param['value'] = value
                        print(f'Successfully written new value {value} in {param["name"]}')
                        return True
                err = 'Не найден параметр в списке'
            else:
                err = f'Текущий параметр из устройства отличается от желаемого {value} <> {data}'
        else:
            err = data
    QMessageBox.critical(window, "Ошибка ", err, QMessageBox.Ok)
    return False


def get_param(address):
    data = 'OK'
    request_iteration = 3
    address = int(address)
    LSB = address & 0xFF
    MSB = ((address & 0xFF00) >> 8)
    # на случай если не удалось с первого раза поймать параметр,
    # делаем ещё request_iteration запросов
    for i in range(request_iteration):
        data = marathon.can_request(current_wheel, current_wheel + 2, [0, 0, 0, 0, LSB, MSB, 0x2B, 0x03])
        if not isinstance(data, str):
            return (data[1] << 8) + data[0]
    QMessageBox.critical(window, "Ошибка ", 'Нет подключения\n' + data, QMessageBox.Ok)
    return False


def get_all_params():
    if get_param(42):
        for param in params_list:
            if str(param['address']) != 'nan':
                if str(param['value']) == 'nan':
                    param['value'] = get_param(address=int(param['address']))
        return True
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


def has_wheel(wheel):
    # запрашиваю тип оси
    err = marathon.can_request(wheel, wheel + 2, [0, 0, 0, 0, 0x23, 0, 0x2B, 0x03])
    if not isinstance(err, str):
        return err[0]
    else:
        return False


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


class ExampleApp(QtWidgets.QMainWindow):
    name_col = 0
    desc_col = 1
    value_col = 3
    combo_col = 999
    unit_col = 3
    record_vmu_params = False

    def __init__(self):
        super().__init__()
        # Это нужно для инициализации нашего дизайна
        uic.loadUi('CANAnalyzer_2.ui', self)
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
        self.hidden_tab = self.Wait_for_read

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
        for par in vmu_params_list:
            value_Item = QTableWidgetItem(str(par['value']))
            value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
            self.vmu_param_table.setItem(row, window.value_col, value_Item)
            row += 1

    def setting_current_wheel(self, item):
        rb_toggled = QApplication.instance().sender()

        if rb_toggled != self.set_front_wheel_rb:
            print('попытка установки передней оси')
            change_current_wheel(2)
            self.set_front_wheel_rb.setChecked(True)
            self.factory_settings_rb.setCheckable(False)
            return True

        elif rb_toggled != self.set_rear_wheel_rb:
            print('попытка установки задней оси')
            change_current_wheel(3)
            self.set_rear_wheel_rb.setChecked(True)
            self.factory_settings_rb.setCheckable(False)
            return True
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

    # надо проверить эту функцию
    def set_slider(self, item):
        slider = QApplication.instance().sender()
        param = slider.objectName()
        item = slider.value()
        value = item / often_used_params[param]['scale']
        address = often_used_params[param]['address']
        label = getattr(self, 'lab_' + param)
        label.setText(str(value) + often_used_params[param]['unit'])
        # надо сделать цикл раза три запихнуть параметр и проверить, если не получилось - предупреждение
        for i in range(marathon.max_iteration):
            if set_param(address, item):
                print('Checked changed value - OK')
                label.setStyleSheet('background-color: green')
                return True
        QMessageBox.critical(window, "Ошибка ", 'Что-то пошло не по плану,\n данные не записались',
                             QMessageBox.Ok)
        label.setStyleSheet('background-color: red')
        return False

    def best_params(self):
        window.groupBox_4.setEnabled(True)
        window.set_current_wheel.setEnabled(True)
        window.byte_order.setEnabled(True)
        self.lb_soft_version.setText('Версия ПО БУРР ' + str(get_param(42)))

        errors = get_param(0)
        errors_str = ''
        print(errors)
        for err_nom, err_str in errors_list.items():
            if errors & err_nom:
                errors_str += err_str + '\n'
        if errors_str == '':
            errors_str = 'Нет ошибок '
        self.tb_errors.setText(errors_str)

        self.set_front_wheel_rb.toggled.disconnect()
        self.set_rear_wheel_rb.toggled.disconnect()
        c_wheel = get_param(0x23)
        print(c_wheel)
        if c_wheel > 1:
            self.factory_settings_rb.setCheckable(False)
            if c_wheel == 2:
                self.set_front_wheel_rb.setChecked(True)
            elif c_wheel == 3:
                self.set_rear_wheel_rb.setChecked(True)
        else:
            self.factory_settings_rb.setChecked(True)

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
        self.params_table.itemChanged.disconnect()
        show_empty_params_list(item, 'params_table')
        self.params_table.itemChanged.connect(self.save_item)

        show_value(self.value_col, item, 'params_table')

    def save_item(self, item):
        table_param = QApplication.instance().sender()
        new_value = item.text()
        if new_value:
            name_param = table_param.item(item.row(), self.name_col).text()
            # если изменили значение параметра
            if item.column() == self.value_col:
                address_param = get_address(name_param)
                if str(address_param) != 'nan':
                    value = check_param(address_param, new_value)
                    if str(value) != 'nan':  # прошёл проверку
                        table_param.item(item.row(), self.value_col).setSelected(False)

                        if set_param(address_param, value):
                            print('Checked changed value - OK')
                            table_param.item(item.row(), self.value_col).setBackground(QColor('green'))
                            return True
                        else:
                            table_param.item(item.row(), self.value_col).setBackground(QColor('red'))
                            table_param.itemChanged.disconnect()
                            table_param.item(item.row(), self.value_col).setText(str(get_param(address_param)))
                            table_param.itemChanged.connect(window.save_item)

                            return False
                        # # ------------ЗДЕСЬ всё не так просто - надо разбираться как действовать при смене рейки
                        # if address_param == 35:  # если произошла смена рейки, нужно поменять адреса
                        #     if self.radioButton.isChecked():
                        #         self.radioButton_2.setChecked()
                        #     else:
                        #         self.radioButton.setChecked()
                        #     rb_clicked()
                        #     return True
                        # return False
                    else:
                        err = "Param isn't in available range - не прошёл проверку check_param"
                else:
                    err = "Can't find this value - не найден адрес параметра"
            elif item.column() == self.desc_col:
                for param in params_list:
                    if name_param == str(param['name']):
                        param['description'] = new_value
                        # здесь должна быть функция по перезаписи исходного файла для дополнения дескрипшн
                        return True
                err = "Can't find this value - в списке параметров этот параметр не найден"
            else:
                err = "It's impossible!!! - действительно херня какая-то , не та колонка таблицы"
        else:
            err = 'Value is empty'
        QMessageBox.critical(window, "Ошибка ", err, QMessageBox.Ok)
        table_param.item(item.row(), item.column()).setSelected(False)
        table_param.item(item.row(), item.column()).setBackground(QColor('red'))

        return False


app = QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

dir_path = str(pathlib.Path.cwd())
# vmu_param_file = 'table_for_params.xlsx'
vmu_param_file = 'table_for_params_forward_wheels.xlsx'
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
# для релиза отключу панель настройки рулевых реек и кнопку выбора файла параметров для КВУ
# window.select_file_vmu_params.setEnabled(False)
# window.burr30.setEnabled(False)
# первое подключение и последующее обновление текущего вида параметров - второе работает не очень
window.pushButton_2.clicked.connect(update_param)
# переключение между задним и передним блоком
window.radioButton.toggled.connect(rb_clicked)
window.radioButton_2.toggled.connect(rb_clicked)
# сохранение всех параметров из текущей рейки в файл
window.pushButton.clicked.connect(save_all_params)

window.list_bookmark.setCurrentRow(0)
#  заполняю все таблицы пустыми параметрами
# заглушка
window.vmu_param_table.itemChanged.connect(check_connection)
show_empty_params_list(bookmark_dict[window.list_bookmark.currentItem().text()], 'params_table')
show_empty_params_list(editable_params_list, 'params_table_2')
show_empty_params_list(vmu_params_list, 'vmu_param_table')
# изменение параметра рейки ведёт к сохранению
window.params_table.itemChanged.connect(window.save_item)
window.params_table_2.itemChanged.connect(window.save_item)

# параметры КВУ только для просмотра, поэтому отключаю их изменение - оно для всех подключается в функции заполнения
# window.vmu_param_table.itemChanged.disconnect()
#  щелчок на списке групп параметров ведёт к выводу этих параметров
window.list_bookmark.itemClicked.connect(window.list_of_params_table)
window.params_table.resizeColumnsToContents()
window.load_file_button.clicked.connect(make_compare_list)
window.load_to_device_button.clicked.connect(write_all_from_file_to_device)

#  часто используемые
# выставляю в нули слайдеры и их метки
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
# изменение текущего блока на противоположный - работает паршиво - нет заводского режима
window.set_front_wheel_rb.toggled.connect(window.setting_current_wheel)
window.set_rear_wheel_rb.toggled.connect(window.setting_current_wheel)
# изменение порядка следования байт - работает тоже паршиво
window.rb_big_endian.toggled.connect(window.set_byte_order)
window.rb_little_endian.toggled.connect(window.set_byte_order)

# главные кнопки для КВУ
window.connect_vmu_btn.clicked.connect(connect_vmu)
window.start_record.clicked.connect(start_btn_pressed)
window.constantly_req_vmu_params.toggled.connect(const_req_vmu_params)
# в окошке с задержкой опроса могут быть только цифры
reg_ex_2 = QRegExp("[0-9]{1,5}")
window.response_time_edit.setValidator(QRegExpValidator(reg_ex_2))
window.response_time_edit.setText('1000')
# window.response_time_edit.textEdited.connect(check_response_time)
window.select_file_vmu_params.clicked.connect(make_vmu_params_list)

# убираю вкладку с ожиданием
window.CAN.removeTab(2)
window.load_to_device_button.hide()
window.show()  # Показываем окно
app.exec_()  # и запускаем приложение
