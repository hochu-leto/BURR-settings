from canlib import canlib, Frame

'''
    canStatus = {
    canOK = 0,
    canERR_PARAM = -1,
    canERR_NOMSG = -2,
    canERR_NOTFOUND = -3,
    canERR_NOMEM = -4,
    canERR_NOCHANNELS = -5,
    canERR_INTERRUPTED = -6,
    canERR_TIMEOUT = -7,
    canERR_NOTINITIALIZED = -8,
    canERR_NOHANDLES = -9,
    canERR_INVHANDLE = -10,
    canERR_INIFILE = -11,
    canERR_DRIVER = -12,
    canERR_TXBUFOFL = -13,
    canERR_RESERVED_1 = -14,
    canERR_HARDWARE = -15,
    canERR_DYNALOAD = -16,
    canERR_DYNALIB = -17,
    canERR_DYNAINIT = -18,
    canERR_NOT_SUPPORTED = -19,
    canERR_RESERVED_5 = -20,
    canERR_RESERVED_6 = -21,
    canERR_RESERVED_2 = -22,
    canERR_DRIVERLOAD = -23,
    canERR_DRIVERFAILED = -24,
    canERR_NOCONFIGMGR = -25,
    canERR_NOCARD = -26,
    canERR_RESERVED_7 = -27,
    canERR_REGISTRY = -28,
    canERR_LICENSE = -29,
    canERR_INTERNAL = -30,
    canERR_NO_ACCESS = -31,
    canERR_NOT_IMPLEMENTED = -32,
    canERR_DEVICE_FILE = -33,
    canERR_HOST_FILE = -34,
    canERR_DISK = -35,
    canERR_CRC = -36,
    canERR_CONFIG = -37,
    canERR_MEMO_FAIL = -38,
    canERR_SCRIPT_FAIL = -39,
    canERR_SCRIPT_WRONG_VERSION = -40,
    canERR_SCRIPT_TXE_CONTAINER_VERSION = -41,
    canERR_SCRIPT_TXE_CONTAINER_FORMAT = -42,
    canERR_BUFFER_TOO_SMALL = -43,
    canERR_IO_WRONG_PIN_TYPE = -44,
    canERR_IO_NOT_CONFIRMED = -45,
    canERR_IO_CONFIG_CHANGED = -46,
    canERR_IO_PENDING = -47,
    canERR_IO_NO_VALID_CONFIG = -48,
    canERR__RESERVED = -49
    }
'''


class Kvaser:
    can_bitrate = {
        125: canlib.Bitrate.BITRATE_125K,
        250: canlib.Bitrate.BITRATE_250K,
        500: canlib.Bitrate.BITRATE_500K
    }

    def __init__(self, channel=0, bit=125):
        self.openFlags = canlib.Open.ACCEPT_VIRTUAL
        self.outputControl = canlib.Driver.NORMAL
        self.can_canal_number = channel  # по умолчанию нулевой канал
        if bit in self.can_bitrate.keys():
            self.bitrate = self.can_bitrate[bit]
        else:
            self.bitrate = canlib.Bitrate.BITRATE_125K  # и скорость 125

        self.wait_time = 1000
        self.max_iteration = 10
        # может, это не совсем верный подход, но я пытаюсь стандартизировать под марафон
        self.ch = self.canal_open()

    def canal_open(self):
        try:
            ch = canlib.openChannel(self.can_canal_number, self.bitrate, self.openFlags)
            ch.setBusOutputControl(self.outputControl)
            ch.busOn()
        except canlib.canError as ex:
            ch = ex
            print(ch)
        return ch

    def close_canal_can(self):
        if isinstance(self.ch, canlib.Channel):
            try:
                self.ch.busOff()
                self.ch.close()
                return ''
            except canlib.canError as ex:
                print(ex)
                return ex
        return self.ch

    def can_write(self, ID: int, data: list):
        if isinstance(data, list):
            frame = Frame(
                id_=ID,
                data=data,
                flags=canlib.MessageFlag.EXT)

            if not isinstance(self.ch, canlib.Channel):
                self.ch = self.canal_open()

            if isinstance(self.ch, canlib.Channel):
                try:
                    self.ch.write(frame)
                    return ''
                except canlib.canError as ex:
                    print(ex)
                    return ex
            else:
                return f'Нет открытого канала {self.ch}'
        else:
            return 'Неправильные данные для передачи. Нужен список'

    def can_read(self, ID: int):
        last_frame_time = int(round(time.time() * 1000))

        if not isinstance(self.ch, canlib.Channel):
            self.ch = self.canal_open()

        if isinstance(self.ch, canlib.Channel):
            while True:
                current_time = int(round(time.time() * 1000))
                if current_time > (last_frame_time + self.wait_time):
                    return f'Истекло время ожидания ответа {self.wait_time}'
                try:
                    frame = self.ch.read()
                except canlib.canError as ex:
                    print(ex)
                    return ex
                if frame.id == ID:
                    return frame.data
        else:
            return f'Нет открытого канала {self.ch}'

    def can_request(self, can_id_req: int, can_id_ans: int, message: list):
        if not isinstance(message, list):
            return 'Неправильные данные для передачи. Нужен список'

        if not isinstance(self.ch, canlib.Channel):
            self.ch = self.canal_open()

        if isinstance(self.ch, canlib.Channel):
            pass
        else:
            return f'Нет открытого канала {self.ch}'

