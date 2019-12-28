from ..simreport import SimulationReport
from abc import ABC

class SIReport(SimulationReport):
    """
    Class for PCB signal integrity report
    """
    def __init__(self, template, interface, proj_num):
        super().__init__(template, proj_num)
        self.__interface = interface
    
    def __str__(self):
        return f"{self.report_type} Report for {self.interface}"

    @property
    def title(self):
        return f"{self.proj_num}\nVerification of Signal Integrity\n{self.interface} [Ver.1.0]"

    @property
    def report_type(self):
        return "SI"

    @property
    def interface(self):
        return self.__interface

    def copy_from_conf(self, conf_tools):
        pass
    
    def clone_temp_slides(self):
        pass