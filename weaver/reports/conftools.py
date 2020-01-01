import re
from .report import Report
from util import COVER_SLIDE, TITLE_NAME, TABLE_COORDS, TOC, com_error


# =======================
# -- Helper Functions --
# =======================

def _set_type(pptx):
    """
    Parses path for report type 
    and verifies if report type is valid
    """
    rep_type = ""
    # Get filename and search for report type
    match = re.search(r"^\w{2}\d{4}.*_(\w{2,3})_", pptx.Name)
    if match:
        rep_type = match.group(1).lower()

        # Verify report type
        if rep_type not in ["emc", "pi", "si"]:
            print(f"ERROR: FILENAME '{pptx.Name}' or PATH '{pptx.FullName}' is not valid")
            raise FilenameError

    return rep_type


# =======================
# -- Class Definition --
# =======================

class ConfirmationTools(Report):
    """
    Class for initial, pre-simulation report
    """
    def __init__(self, pptx):
        super().__init__(pptx)
        # Regex project number from title
        self.__proj_num = re.search(r"(^\w{2}\d{4})", self.title).group(1)[:] 
        self.__toc = None
        self.__type = _set_type(self.pptx)

    @property
    def title(self):
        """
        Fetches title from cover slide
        """
        # Pull title from cover slide
        return self.pptx.Slides(COVER_SLIDE).\
               Shapes(TITLE_NAME).TextFrame.TextRange.Text[:] 
    
    @property
    def type(self):
        """
        Returns ConfirmationTools type
        """
        return self.__type

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
        # TODO: unmemoize if not needed
        if not self.__toc:
            toc = self._get_table(self.pptx.Slides(TOC).Shapes)
            # To be populated with slide nums
            toc_dict = { "sim_target": None }

            # Add depending on type
            if self.type == "si":
                toc_dict["topology"] = None 
                toc_dict["eye_mask_judgement"] = None
            elif self.type == "pi":
                toc_dict["curr_consumption"] = None
                toc_dict["voltage_margin"] = None
                toc_dict["appendix"] = None
            # TODO
            # elif self.type == "emc":
            #     toc_dict[""]

            row = 2 # Starting y coord of table traversal
            while True:
                section_name = toc.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
                section_name = section_name.lower()
                # Only contents in section 2 is of interest
                try:
                    if re.search(r"^\s*\d?\.?\d?\w+", section_name):
                        for key in toc_dict.keys():
                            target = key.split("_")[-1] if key.find("_") > -1 else key
                            if self.type == "si":
                                section_name = section_name.split("&")[0].strip()
                            if section_name.endswith(target):
                                slide_num = toc.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                                toc_dict[key] = slide_num
                    # Check if end of TOC in order to end loop
                    elif section_name == "":
                        break
                    # Move down TOC
                    row += 1 
                # In case pointer has reached end of TOC
                except com_error:
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
                    # Keep returned data structures consistent by keeping values as list type
                    num = [ int(slide_nums) ] * 2
                    toc_dict[section] = num

            print()
            print("Loaded page numbers of the following sections:")
            for k in toc_dict:
                print(f"  {k.upper()}: {toc_dict[k][0]} - {toc_dict[k][1]}")
            print()
            self.__toc = toc_dict

        return self.__toc


class FilenameError(Exception):
    pass