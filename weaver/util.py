import os
# from .reports.meta import Interface, Signal
from pywintypes import com_error

# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** GLOBAL CONSTANTS ****
# //////////////////////////////

# Path to textfile listing template paths
TXT_PATH = r"E:\vs_code\takehome\Weaver\paths_to_templates.txt"

# For cover slide
COVER_SLIDE = 1 # Indexing starts at 1 for COM Objects
TABLE_COORDS = {
    "preparers": (1,2),
    "reviewers": (2,2),
    # "approvers": (3,1) 
}

# Slide index of table of contents
# and executive summary
TOC = 3
EXEC_SUMM = 4

# Values to verify shape identity
MSOTRUE = -1
TITLE_NAME = "Rectangle 26" 
REP_SLIDE_TITLE = "Title 6" # TODO: make consistent between slides
DATE_NAME = u"テキスト プレースホルダー 10"


def _parse_if_name(shapes):
    """
    Searches through Slide.Shapes for Shape with Text;
    Returns Interface.name if found based on pattern
    """
    if_name = ""
    last_title = ""
    for shape in shapes:
        if shape.HasTextFrame == MSOTRUE:
            text = shape.TextFrame.TextRange.Text[:].lower()
            curr_title = text[:]
            if last_title != curr_title: 
                last_title = curr_title
                if text.find("target & condition") > -1:
                    if text.find(":") > -1:
                        # Displace pointer to the right by 1 and strip spaces
                        if_name = text[text.find(":")+1:].strip()
                    # # In case full-size colon used
                    # except:
                    #     tar_index = text.find("：")
                    
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
        except com_error:
            break

    # Use simulation directory for ibis models if not found in confirmation tools
    if sim_dir and not signal.driver.ibis_model and not signal.receiver.ibis_model:
        for device in [ signal.driver, signal.receiver ]:
            device.ibis_model =  _get_ibis_models(interface.name, signal.name, sim_dir)
        
    return signal


def _set_signal(table):
    """
    Set signal features based on target and frequency table
    """
    signal = Signal()
    row = 2
    while True:
        # Set name
        try:
            signal_group = table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
            index = signal_group.find(":")
            signal.name = signal_group[index+1:] if index > -1 else signal_group
            if index > -1: 
                signal.type = signal_group[:index].strip()
            # print(signal.name)

            # Set frequency
            freq_str = table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
            try:
                signal.frequency = freq_str.split()[0].strip(), freq_str.split()[1].strip()
            except IndexError:
                signal.frequency = None

            # Set driver / receiver
            trans_line = table.Cell(row, 3).Shape.TextFrame.TextRange.Text[:]
            signal.driver.ref_num = trans_line.split("~")[0].strip()
            signal.receiver.ref_num = trans_line.split("~")[1].strip()

            # Set PVT value
            signal.pvt = table.Cell(row, 5).Shape.TextFrame.TextRange.Text[:].split("/")
            signal.pvt = [ item.strip() for item in signal.pvt ]
            
            yield signal
            row += 1

        except com_error:
            break
    

def _read_interface(slide, if_name, sim_dir):
    """
    Factory function for Interface instances with all fields filled in 
    based on data found on the current Slide
    """
    interface = Interface(if_name)

    tar_and_freq_table, ic_model_table = _get_if_tables(slide.Shapes)
    if tar_and_freq_table and ic_model_table:
        for signal in _set_signal(tar_and_freq_table): 
            interface.signals.append(signal)
        print()
        for i, signal in enumerate(interface.signals):
            interface.signals[i] = _set_signal_devices(interface, signal, ic_model_table, sim_dir)
            print(f"Loaded the following data for")
            print(f"{signal.name}:")
            print(f"DRIVER: {signal.driver.ref_num}")
            print(f"RECEIVER: {signal.receiver.ref_num}\n")
        print(f"TOTAL SIGNALS in")
        print(f"{interface.name}: {len(interface.signals)}")
        print()
        return interface


def get_interfaces(conf_tools, sim_dir):
    toc = conf_tools.get_toc()
    start, end = toc["sim_target"][0], toc["sim_target"][1]
    for i in range(start, end + 1):
        # last_title = ""
        slide = conf_tools.pptx.Slides(i)
        if_name = _parse_if_name(slide.Shapes)
        interface = _read_interface(slide, if_name, sim_dir)
        if interface:
            yield interface

from abc import ABC
class Interface():
    def __init__(self, name):
        self.__name = name.upper()
        self.signals = list()

    @property
    def name(self):
        return self.__name[:]


class Device(ABC):
    def __init__(self):
        self.ref_num = str()
        self.part_name = str()
        self.ibis_model = str()
        self.buffer_model = str()

    # def ref_num(self):
    #     return self.__ref_num[:]
    
    # def part_name(self):
    #     return self.__part_name[:]
    
    # def ibis_model(self):
    #     return self.__ibis_model[:]

    # def buffer_model(self):
    #     return self.__buffer_model[:]


class Driver(Device):
    def __init__(self):
        super().__init__()

class Receiver(Device):
    def __init__(self):
        super().__init__()

class Signal():
    def __init__(self):
        self.type = str()
        self.name = str()
        self.__driver = Driver()
        self.__receiver = Receiver()
        self.pvt = str()
        self.frequency = None

    @property
    def driver(self):
        return self.__driver
    
    @property
    def receiver(self):
        return self.__receiver