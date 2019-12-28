import re, os
from .report import Report
from .util import Interface, Signal
from ..consts import COVER_SLIDE, TITLE_NAME, TABLE_COORDS, TOC, MSOTRUE


# =======================
# -- Helper Functions --
# =======================

def _parse_if_name(shapes):
    """
    Searches through Slide.Shapes for Shape with Text;
    Returns Interface.name if found based on pattern
    """
    if_name = ""
    for shape in shapes:
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
                    # Displace pointer to the right by 1 and strip spaces
                    if_name = text[tar_index+1:].strip()
                    
    return if_name


def _get_if_tables(shapes):
    """
    Returns pointers to the Target and Frequency and IC Model tables
    """
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


def _get_ibis_models(if_name, sig_name, sim_dir):
    """
    Returns a str to be set as the ibis_model of a Signal.Device
    """
    # Path requires particular directory structure
    signal_path = os.path.join(sim_dir, if_name, sig_name)
    if not os.path.exists(signal_path):
        print(f"Could not find {signal_path}:\n  \
                Skipping addition of IBIS Model info for {sig_name} in {if_name}")
    # Look through folders or current folder for IBIS files
    # Let user choose during report editing which is correct
    ibis_str = ""
    in_folder = False # Flag for whether .ibs in signal_path or folder therewithin
    for item in os.listdir(signal_path):
        ext = os.path.splitext(item)[1] # Get file ext
        # Found a file -- check if .ibs
        if ext:
            if ext == ".ibs":
                in_folder = True
                ibis_str += item + " " # Add space for additional ibis file
        # If .ibs not yet found in current path, check a level deeper
        elif not in_folder:
            for subitem in os.listdir(os.path.join(signal_path, item)):
                if os.path.splitext(subitem)[1] == ".ibs":
                    ibis_str += subitem + " "
    if not ibis_str:
        print(f"Could not find IBIS Models for {sig_name} in {if_name}")

    return ibis_str


def _set_signal_devices(interface, signal, table, sim_dir):
    """
    Sets the Driver and Receiver of an input signal
    """
    row = 1
    while True:
        # Go through every column and fill in Signal fields accordingly
        try:
            ref_num = table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
            ic_model = table.Cell(row, 4).Shape.TextFrame.TextRange.Text[:]

            if signal.driver.ref_num == ref_num:
                signal.driver.part_name = table.Cell(row, 3).\
                                            Shape.TextFrame.TextRange.Text[:].split()[1]
                signal.driver.ibis_model = None if ic_model.find("?") > -1 else ic_model
                
            elif signal.receiver.ref_num == ref_num:
                signal.receiver.part_name = table.Cell(row, 3).\
                                            Shape.TextFrame.TextRange.Text[:].split()[1]
                signal.receiver.ibis_model = None if ic_model.find("?") > -1 else ic_model
                
            row += 1

        # Reached end of table
        except IndexError:
            break

    # Use simulation directory for ibis models if not found in confirmation tools
    if sim_dir and not signal.driver.ibis_model and not signal.receiver.ibis_model:
        for device in [ signal.driver, signal.receiver ]:
            device.ibis_model =  _get_ibis_models(interface.name, signal.name, sim_dir)
        
    return signal


def _set_signal(signal, table, row):
    """
    Set signal features based on target and frequency table
    """
    # Set name
    signal_group = table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
    index = signal_group.find(":")
    signal.name = signal_group[index+1:] if index > -1 else signal_group
    if index > -1: 
        signal.type = signal_group[:index]

    # Set frequency
    freq_str = table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
    signal.frequency = freq_str.split()[0], freq_str.split()[1]

    # Set driver / receiver
    trans_line = table.Cell(row, 3).Shape.TextFrame.TextRange.Text[:]
    signal.driver.ref_num = trans_line.split("~")[0]
    signal.receiver.ref_num = trans_line.split("~")[1]

    # Set PVT value
    signal.pvt = table.Cell(row, 5).Shape.TextFrame.TextRange.Text[:]
    
    return signal


def _read_interface(slide, if_name, sim_dir):
    """
    Factory function for Interface instances with all fields filled in 
    based on data found on the current Slide
    """
    interface = Interface(if_name)

    tar_and_freq_table, ic_model_table = _get_if_tables(slide.Shapes)

    while True:
        row = 2
        try:
            # Instantiate and set signal
            signal = Signal()
            interface.signals.append(_set_signal(signal, tar_and_freq_table, row))

        except IndexError:
            break
    
    for i, signal in enumerate(interface.signals):
        interface.signals[i] = _set_signal_devices(interface, signal, ic_model_table, sim_dir)

    return interface



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
        if not self.__toc:
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

            self.__toc = toc_dict
        
        return self.__toc

    def get_interfaces(self, sim_dir=""):
        toc = self.get_toc()
        start, end = toc["sim_targets"][0], toc["sim_targets"][1]
        for i in range(start, end + 1):
            last_title = ""
            slide = self.pptx.Slides(i)
            if_name = _parse_if_name(slide.Shapes)
            interface = _read_interface(slide, if_name, sim_dir)
            yield interface