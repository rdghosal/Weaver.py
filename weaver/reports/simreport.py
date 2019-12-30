import os
from datetime import date
from .report import Report
from ..util import COVER_SLIDE, TITLE_NAME, DATE_NAME, MSOTRUE, TABLE_COORDS


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

    def _get_date(self):
        """
        Receives isoformat date from user 
        and returns the date formatted according to report standards
        """
        date_str = ""
        while True:
            # Instructions for user input
            prompt = """
            Input report date as follows:

            yyyy-MM-dd

            where:
            yyyy -> year
            MM -> month
            dd -> date
            """
            # Check if instructions were followed
            try:
                date_str = date.fromisoformat(input(prompt))
                break
            except ValueError: 
                continue

        # Formats as e.g. 07 Feb. 2001, 15 Nov. 1753 etc.
        return f"{date_str.strftime('%d %b. %Y')}" 

    def _make_cover(self, conf_tools):
        """
        Sets first slide of report from args and user input
        """
        # Copy cover slide unto Clipboard
        # and paste so as to make it the first slide in the report
        conf_tools.pptx.Slide(COVER_SLIDE).Copy() 
        self.pptx.Slides.Paste(COVER_SLIDE) 

        # Grab the pasted cover slide
        # and iterate over its shapes in order to replace their contents
        cover = self.pptx.Slide(COVER_SLIDE)
        for shape in cover.Shapes:
            if shape.Name == TITLE_NAME:
                shape.TextFrame.TextRange.Text = self.title[:]
            elif shape.Name == DATE_NAME:
                shape.TextFrame.TextRange.Text = self._get_date()
            elif shape.HasTable == MSOTRUE:
                conf_creators = conf_tools.get_creators()
                # Match table coordinates with creator keys and insert values of latter
                for group, coords in TABLE_COORDS.items():
                    cover.Shapes.Table.Cell(coords[0], coords[1]).Shape.TextFrame.TextRange.Text = conf_creators[group][:]

        print(f"Cover slide generated for {self.title}.")
    
    def _copy_slides(self, conf_tools):
        raise NotImplementedError

    def _save_report(self):
        """
        Gets filename from user and closes report after saving
        """
        filename = ""
        path = ""
        while True:
            filename = input(f"Input filename to save the report {self.title}: ")
            path = input("Input path to save report: ") # TODO: develop algorithm to fix name
            if filename and os.path.isabs(path):
                break

        self.pptx.SaveAs(os.path.join(path, filename))
        self.pptx.Close()
        print(f"{filename} saved in {path}.")

    def build_pptx(self, conf_tools):
        raise NotImplementedError