from ..simreport import SimulationReport


class PIReport(SimulationReport):
    """
    Class for PCB power integrity report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)
        self.__net_names = []

    def __str__(self):
        pass

    @property
    def title(self):
        return f"{self.proj_num}\nVerification of Power Integrity by PI Simulation [Ver.1.0]"

    @property
    def report_type(self):
        return "PI"
    
    @property
    def net_names(self):
        return self.__net_names

    @net_names.setter
    def net_names(self, value):
        self.__net_names = value