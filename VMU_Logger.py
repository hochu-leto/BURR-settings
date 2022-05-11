"""
Цель этой ветки - создать  программу для просмотра/записи параметров из кву цикл + и TTC
- никакого управления
- никаких плюшек типа выбора параметров
- выбор только файла с набором параметров и их опрос

- всё, что относится к БУРР-30 - СНОСИМ НАХРЕН!!!!!!
- для него есть свой проект на гитлабе эвокарго
--------------------------------------------критично------------------------------
--------------------------------------------хотелки-------------------------
- сделать поиск параметра по описанию и названию
- задел под парсинг файла с настройками от рткона
-- индекс и саб_индекс вместо адресов кву
--- исправить проверку на адрес, и если есть индекс и саб_индекс, брать их за основу и запихивать в адрес
-- скале_валуе - если есть, и не ноль, то запихиваем в скале))) - если это когда-то понадобится
- ограничить число записываемых параметров не более 30 штук, если в файле больше - брать только первые 50,
 либо предлагать выбрать другой файл
---------------------------------------------непонятки----------------------
- не подключается когда подключили марафон или закрыли рткон или канвайс - тут вообще
какая-то непонятная фигня, есть подозрения, что это особенность работы марафона - без перезагрузки он периодически
отваливается . Выход - переход на квайзер и кан-хакер
--------------------------------------------исправил--------------------------------
"""
import ctypes
import datetime
import pathlib

from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, pyqtSlot, QRegExp
from PyQt5.QtGui import QColor, QRegExpValidator, QFont, QIcon
from dll_power import CANMarathon
from PyQt5 import uic
from PyQt5.QtWidgets import QTableWidgetItem, QApplication, QMessageBox, QFileDialog, QMainWindow
import pandas as pandas

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


class ExampleApp(QMainWindow):
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
        for par in vmu_params_list:
            value_Item = QTableWidgetItem(str(par['value']))
            value_Item.setFlags(value_Item.flags() & ~Qt.ItemIsEditable)
            self.vmu_param_table.setItem(row, window.value_col, value_Item)
            row += 1


app = QApplication([])
window = ExampleApp()  # Создаём объект класса ExampleApp

dir_path = str(pathlib.Path.cwd())
# vmu_param_file = 'table_for_params.xlsx'
vmu_param_file = 'wheels&RPM.xlsx'
# vmu_param_file = 'table_for_params_forward_wheels.xlsx'
vmu_params_list = fill_vmu_list(pathlib.Path(dir_path, 'Tables', vmu_param_file))
# заполняю дату с адресами параметров из списка, который задаётся в файле
req_list = feel_req_list(vmu_params_list)

marathon = CANMarathon()
show_empty_params_list(vmu_params_list, 'vmu_param_table')
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
