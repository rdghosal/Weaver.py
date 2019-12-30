from .report import Report


class SimulationReport(Report):
    """
    Base class for simulation reports
    """
    __rep_types = ["si", "pi", "emc", "thermal"]

    def __init__(self, pptx_template, proj_num):
        super().__init__(pptx_template)
        self.__proj_num = proj_num
        self.__curr_slide = 1

    @staticmethod
    def report_types():
        return SimulationReport.__rep_types

    @property
    def report_type(self):
        raise NotImplementedError 

    def build_slides(self, conf_tools):
        raise NotImplementedError