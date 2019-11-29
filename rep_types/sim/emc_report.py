from ..simreport import SimulationReport


class EMCReport(SimulationReport):
    """
    Class for PCB EMC report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)

    def __str__(self):
        pass

    @property
    def title(self):
        return f"{self.proj_num}\nEMC (Power Resonance) Simulation [Ver.1.0]"
    
    @property
    def report_type(self):
        return "EMC"