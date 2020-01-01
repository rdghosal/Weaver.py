from util import MSOTRUE


class Report():
    """
    Base class for simulation report
    """
    def __init__(self, pptx_obj):
        self.__pptx = pptx_obj
        self.__title = ""
        self.__proj_num = ""
    
    @property
    def pptx(self):
        """
        Returns instance of PowerPoint COM Object associated with report
        """
        return self.__pptx
    
    @property
    def proj_num(self):
        """
        Returns project number of report
        """
        return self.__proj_num

    @property
    def title(self):
        raise NotImplementedError

    def _get_table(self, shapes): 
        """
        Iterates over a Slide's collection of Shapes 
        and returns first shape found to have a Table object
        """
        for shape in shapes:
            if shape.HasTable == MSOTRUE:
                return shape.Table
        return None 

