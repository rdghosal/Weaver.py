import re
from .report import Report
from .util import Interface, Signal
from ..consts import COVER_SLIDE, TITLE_NAME, TABLE_COORDS, TOC, MSOTRUE


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

    def get_interfaces(self, conf_tools):
        toc = conf_tools.get_toc()
        start, end = toc["sim_targets"][0], toc["sim_targets"][1]
        for i in range(start, end + 1):
            last_title = ""
            slide = self.pptx.Slides(i)
            for shape in slide.Shapes:
                if shape.HasTextFrame == MSOTRUE:
                    text = shape.TextFrame.TextRange.Text[:].lower()
                    curr_title = text[:]
                    if last_title != curr_title: 
                        last_title = curr_title
                        if text.find("target") > -1:
                            tar_index = text.index(":")
                            # In case full-size colon used
                            if not tar_index:
                                tar_index = text.find("：")
                            # Displace pointer to the right by 1
                            if_name = text[tar_index+1:].strip()
                            yield _read_interface(slide, if_name)


def _read_interface(slide, if_name):
    """
    Factory function for Interface instances with all fields filled in 
    based on data found on the current Slide
    """
    interface = Interface(if_name)

    def get_if_tables(shapes):
        """Returns pointers to the Target and Frequency and IC Model tables"""
        tar_and_freq_table = None
        ic_model_table = None
        for s in shapes:
            table_count = 0
            if s.HasTable == MSOTRUE:
                table_count += 1
                first_header_name = s.Table.Cell(1,1).Shape.TextFrame.TextRange.Text[:]
                if first_header_name == "Signal Group":
                    tar_and_freq_table = s.Table
                elif first_header_name == "Reference":
                    ic_model_table = s.Table
                else:
                    print(f"Found table with header name '{first_header_name}'")
            if table_count == 2:
                break

        return tar_and_freq_table, ic_model_table

    tar_and_freq_table, ic_model_table = get_if_tables(slide.Shapes)

    while True:
        row = 2
        try:
            signal = Signal()
            
            # Set name
            signal_group = tar_and_freq_table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
            index = signal_group.find(":")
            signal.name = signal_group[index+1:] if index > -1 else signal_group
            if index > -1: 
                signal.type = signal_group[:index]

            # Set frequency
            freq_str = tar_and_freq_table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
            signal.frequency = freq_str.split()[0], freq_str.split()[1]

            # Set driver / receiver
            trans_line = tar_and_freq_table.Cell(row, 3).Shape.TextFrame.TextRange.Text[:]
            signal.driver.ref_num = trans_line.split("~")[0]
            signal.receiver.ref_num = trans_line.split("~")[1]

            # Set PVT value
            signal.pvt = tar_and_freq_table.Cell(row, 5).Shape.TextFrame.TextRange.Text[:]
            
            interface.signals.append(signal)

        except IndexError:
            break
    
    for i, signal in enumerate(interface.signals):
        row = 1
        while True:
            # Go through every column and fill in Signal fields accordingly
            try:
                ref_num = ic_model_table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
                ic_model = ic_model_table.Cell(row, 4).Shape.TextFrame.TextRange.Text[:]

                if signal.driver.ref_num == ref_num:
                    signal.driver.part_name = ic_model_table.Cell(row, 3).\
                                              Shape.TextFrame.TextRange.Text[:].split()[1]
                    signal.driver.ibis_model = None if ic_model.find("?") > -1 else ic_model
                    
                elif signal.receiver.ref_num == ref_num:
                    signal.receiver.part_name = ic_model_table.Cell(row, 3).\
                                                Shape.TextFrame.TextRange.Text[:].split()[1]
                    signal.receiver.ibis_model = None if ic_model.find("?") > -1 else ic_model
                    
                row += 1

            # Reached end of table
            except IndexError:
                interface.signals[i] = signal
                break

    return interface