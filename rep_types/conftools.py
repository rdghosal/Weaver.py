import re
from .report import Report
from ..consts import COVER_SLIDE, TITLE_NAME, TABLE_COORDS, TOC

class ConfirmationTools(Report):
    """
    Class for initial, pre-simulation report
    """
    def __init__(self, pptx):
        super().__init__(pptx)
        # Regex project number from title
        self.__proj_num = re.search(r"(^\w{2}\d{4})", self.title).group(1)[:] 

    @property
    def title(self):
        """
        Fetches title from cover slide
        """
        # Pull title from cover slide
        return self.pptx.Slides(COVER_SLIDE).\
               Shapes(TITLE_NAME).TextFrame.TextRange.Text[:] 

    def get_creators(self):
        """
        Gets list of authors, reviewers, and approvers 
        from Confirmation Tools object
        """
        creators = {
            "preparers": "",
            "reviewers": "",
            # "approvers": ""
        }

        for party, coords in TABLE_COORDS.items():
            if party in creators.keys():
                creators_table = self._get_table(self.pptx.Slides(COVER_SLIDE).Shapes)
                creators[party] = creators_table.\
                                  Cell(coords[0], coords[1]).Shape.TextFrame.TextRange.Text[:]

        return creators

    def get_toc(self):
        """
        Returns dict of section->slide_num(s) for sections of interest
        """
        toc = self._get_table(self.pptx.Slides(TOC).Shapes)
        # To be populated with slide nums
        toc_dict = {
            "sim_targets": None,
            "eye_masks": None,
            "topology": None
        }

        row = 2 # Starting y coord of table traversal
        while True:
            section_name = toc.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
            section_name = section_name.lower()
            # Only contents in section 2 is of interest
            try:
                if re.search(r"^\s*2\.", section_name):
                    if section_name.find("target") > -1:
                        slide_num = toc.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                        toc_dict["sim_targets"] = slide_num
                    elif section_name.find("mask") > -1:
                        slide_num = toc.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                        toc_dict["eye_masks"] = slide_num
                    elif section_name.find("topology") > -1:
                        slide_num = toc.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                        toc_dict["topology"] = slide_num
                # Check if end of TOC in order to end loop
                elif section_name == "":
                    break
                # Move down TOC
                row += 1 
            # In case pointer has reached end of TOC
            except IndexError:
                break

        # Convert str slide_nums to int for slide indexing
        for section, slide_nums in toc_dict.items():
            # Check if range of slide_nums
            # In case of hyphen type -
            if slide_nums.find("-") > -1:
                slide_nums = slide_nums.split("-")
                toc_dict[section] = [ int(num) for num in slide_nums ]
            # In case of hyphen type ―    
            elif slide_nums.find("\u2013") > -1:
                slide_nums = slide_nums.split("\u2013") 
                toc_dict[section] = [ int(num) for num in slide_nums ]
            # If single number
            else:
                print(slide_nums)
                # Keep returned data structures consistent by keeping values as list type
                toc_dict[section] = toc_dict.get(section, list(int(slide_nums)))
        
        return toc_dict

