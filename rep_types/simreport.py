from .report import Report


class SimulationReport(Report):
    """
    Base class for simulation reports
    """
    __rep_types = ["si", "pi", "emc", "thermal"]

    def __init__(self, pptx_template, proj_num):
        super().__init__(pptx_template)
        self.__proj_num = proj_num

    @staticmethod
    def report_types():
        return SimulationReport.__rep_types

    @property
    def report_type(self):
        raise NotImplementedError 

    def read_toc(self):
        raise NotImplementedError

    # @abstractmethod
    # def __str__(self):
    #     return NotImplementedError

    