from ..simreport import SimulationReport


class ThermalReport(SimulationReport):
    """
    Class for PCB thermal report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)

    def __str__(self):
        pass