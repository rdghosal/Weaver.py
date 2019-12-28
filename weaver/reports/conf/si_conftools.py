import os
from .. import ConfirmationTools, Interface, Signal





class SIConfirmationTools(ConfirmationTools):

    def __init__(self, pptx, type_):
        super().__init__(pptx)
        self.__type = type_

    # def get_interfaces(self, sim_dir=""):