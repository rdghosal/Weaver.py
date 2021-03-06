import os
from time import sleep
from datetime import date
from .report import Report
from util import COVER_SLIDE, TITLE_NAME, DATE_NAME, MSOTRUE, TABLE_COORDS


class SimulationReport(Report):
    """
    Base class for simulation reports
    """
    __rep_types = ["si", "pi", "emc", "thermal"]

    def __init__(self, pptx_template, proj_num):
        super().__init__(pptx_template)
        self.__proj_num = proj_num
        self._curr_slide = 1

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
            prompt = """Input report date as follows: yyyy-MM-dd\nWhere:\n  yyyy -> year\n  MM -> month\n  dd -> date\n\nDate: """
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
        conf_tools.pptx.Slides(COVER_SLIDE).Copy() 
        sleep(.25)
        self.pptx.Slides.Paste(COVER_SLIDE) 

        # Grab the pasted cover slide
        # and iterate over its shapes in order to replace their contents
        cover = self.pptx.Slides(COVER_SLIDE)
        for shape in cover.Shapes:
            if shape.Name == TITLE_NAME:
                shape.TextFrame.TextRange.Text = self.title[:]
            elif shape.Name == DATE_NAME:
                shape.TextFrame.TextRange.Text = self._get_date()
            elif shape.HasTable == MSOTRUE:
                conf_creators = conf_tools.get_creators()
                # Match table coordinates with creator keys and insert values of latter
                for group, coords in TABLE_COORDS.items():
                    shape.Table.Cell(coords[0], coords[1]).Shape.TextFrame.TextRange.Text = conf_creators[group][:]

        title = " ".join(self.title[:].split("\n")).strip()
        print(f"Cover slide generated for {title}.\n")
    
    def _copy_slides(self, conf_tools):
        """
        Copies over slides of interest 
        from confirmation tools to report
        """
        toc = conf_tools.get_toc()
        # Read each slide according to toc
        for section in toc:
            for slide_num in range(toc[section][0], toc[section][1] + 1):
                # Copy current slide unto Clipboard
                conf_tools.pptx.Slides(slide_num).Copy()
                sleep(.25)
                # Paste slide into the same position of report if possible;
                # otherwise, append to end
                pos = slide_num - 1 if slide_num <= len(self.pptx.Slides) else ""
                self.pptx.Slides.Paste(pos)
                self._curr_slide += 1
    
    def _build_slides(self):
        raise NotImplementedError

    def _save_report(self):
        """
        Gets filename from user and closes report after saving
        """
        filename = ""
        path = ""
        while True:
            title = " ".join(self.title[:].split("\n"))
            filename = input(f"Input filename to save the report {title}:\n")
            path = input("Input path to save report: ") # TODO: develop algorithm to fix name
            if filename and os.path.isabs(path):
                if os.path.exists(os.path.join(path, filename)):
                    print("ERROR: File of specified name already exists.")
                    return
                if not os.path.exists(path): 
                    os.mkdir(path)
                if not filename.endswith(".pptx"): filename += ".pptx"
                break

        self.pptx.SaveAs(os.path.join(path, filename))
        self.pptx.Close()
        print(f"{filename} saved in {path}.")

    def build_pptx(self, conf_tools):
        raise NotImplementedError