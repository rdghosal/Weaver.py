from abc import ABC

class Interface():
    def __init__(self, name):
        self.__name = name
        self.signals = list()

    @property
    def name(self):
        return self.__name[:]


class Device(ABC):
    def __init__(self):
        self.ref_num = str()
        self.part_name = str()
        self.ibis_model = str()
        self.buffer_model = str()

    # def ref_num(self):
    #     return self.__ref_num[:]
    
    # def part_name(self):
    #     return self.__part_name[:]
    
    # def ibis_model(self):
    #     return self.__ibis_model[:]

    # def buffer_model(self):
    #     return self.__buffer_model[:]


class Driver(Device):
    def __init__(self):
        super().__init__()

class Receiver(Device):
    def __init__(self):
        super().__init__()

class Signal():
    def __init__(self):
        self.type = str()
        self.name = str()
        self.__driver = Driver()
        self.__receiver = Receiver()
        self.pvt = str()
        self.frequency = None

    @property
    def driver(self):
        return self.__driver
    
    @property
    def receiver(self):
        return self.__receiver