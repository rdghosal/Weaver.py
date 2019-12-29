import os
import win32com.client as win32
from pywintypes import com_error

# from time import sleep
# from abc import ABC, abstractmethod
from datetime import date
from .util import MSOTRUE, TXT_PATH, COVER_SLIDE, TITLE_NAME, DATE_NAME, TABLE_COORDS
from .reports import ConfirmationTools, Signal, Interface
from .reports.sim import SIReport, PIReport, EMCReport



# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////

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
                signal.type = signal_group[:index]
            # print(signal.name)

            # Set frequency
            freq_str = table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
            try:
                signal.frequency = freq_str.split()[0], freq_str.split()[1]
            except IndexError:
                signal.frequency = None

            # Set driver / receiver
            trans_line = table.Cell(row, 3).Shape.TextFrame.TextRange.Text[:]
            signal.driver.ref_num = trans_line.split("~")[0]
            signal.receiver.ref_num = trans_line.split("~")[1]

            # Set PVT value
            signal.pvt = table.Cell(row, 5).Shape.TextFrame.TextRange.Text[:]
            
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
        for i, signal in enumerate(interface.signals):
            interface.signals[i] = _set_signal_devices(interface, signal, ic_model_table, sim_dir)

        return interface


def get_interfaces(conf_tools, sim_dir):
    toc = conf_tools.get_toc()
    start, end = toc["sim_targets"][0], toc["sim_targets"][1]
    for i in range(start, end + 1):
        # last_title = ""
        slide = conf_tools.pptx.Slides(i)
        if_name = _parse_if_name(slide.Shapes)
        interface = _read_interface(slide, if_name, sim_dir)
        if interface:
            yield interface


def _load_template_paths(file_path):
    """
    Fetches template paths and returns dict mapping report type to template
    """
    # dict to be populated
    templates = {
        "si": "",
        "pi": "",
        "emc": "",
        "thermal": "",
    }

    # Fetch paths from text file in same directory
    with open(file_path, "r") as f:
        for rep_type in templates:
            # Iterate over each line prefixed with template key and = (e.g. "si=")
            for line in f.readlines():
                if line.startswith(rep_type):
                    start = line.index("=")
                    # Omit delimiter and copy remaining str into dict
                    templates[rep_type] = line[start+1:] 
    
    return templates


def init_reports(conf_tools, sim_dir=""):
    """
    Initializes and returns Report based on user input and template
    """
    templates = _load_template_paths(TXT_PATH)
    reports = None
    proj_num = conf_tools.proj_num[:]
    rep_type = conf_tools.type

    # Instantiate report based on user input
    if rep_type == "si":
        reports = [ SIReport(templates[rep_type], interface, proj_num) for interface in get_interfaces(conf_tools, sim_dir) ]
    elif rep_type == "pi":
        reports = PIReport(templates[rep_type], proj_num)
    elif rep_type == "emc":
        reports = EMCReport(templates[rep_type], proj_num)

    # Ensure returned object is of consistent data structure
    if not isinstance(reports, list):
        reports = list(reports)

    return reports


def _get_date():
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


def make_cover(conf_tools, report):
    """
    Sets first slide of report from args and user input
    """
    # Copy cover slide unto Clipboard
    # and paste so as to make it the first slide in the report
    conf_tools.pptx.Slide(COVER_SLIDE).Copy() 
    report.pptx.Slides.Paste(COVER_SLIDE) 

    # Grab the pasted cover slide
    # and iterate over its shapes in order to replace their contents
    cover = report.pptx.Slide(COVER_SLIDE)
    for shape in cover.Shapes:
        if shape.Name == TITLE_NAME:
            shape.TextFrame.TextRange.Text = report.title[:]
        elif shape.Name == DATE_NAME:
            shape.TextFrame.TextRange.Text = _get_date()
        elif shape.HasTable == MSOTRUE:
            conf_creators = conf_tools.get_creators()
            # Match table coordinates with creator keys and insert values of latter
            for group, coords in TABLE_COORDS.items():
                cover.Shapes.Table.Cell(coords[0], coords[1]).Shape.TextFrame.TextRange.Text = conf_creators[group][:]

    print(f"Cover slide generated for {report.title}.")


def copy_slides(conf_tools, report):
    """
    Copies over slides of interest 
    from confirmation tools to report
    """
    toc = conf_tools.get_toc()
    # Read each slide according to toc
    for section in toc:
        for slide_nums in toc[section]:
            # Avoids copying pages listed in sim_targets 
            if section == "sim_targets": 
                continue
            for slide_num in slide_nums:
                # Copy current slide unto Clipboard
                conf_tools.pptx.Slides(slide_num).Copy()
                
                # Paste slide into the same position of report if possible;
                # otherwise, append to end
                pos = slide_num if slide_num <= len(report.Slides) else ""
                report.Slides.Paste(pos)


def save_report(report):
    """
    Gets filename from user and closes report after saving
    """
    filename = ""
    path = ""
    while True:
        filename = input(f"Input filename to save the report {report.title}: ")
        path = input("Input path to save report: ") # TODO: develop algorithm to fix name
        if filename and os.path.isabs(path):
            break

    report.pptx.SaveAs(os.path.join(path, filename))
    report.pptx.Close()
    print(f"{filename} saved in {path}.")


def weave_reports(conf_path, sim_dir):
    """
    Generate reports based on input confirmation tools and indicated type
    """
    # Start PowerPoint process
    PowerPoint = win32.gencache.EnsureDispatch("PowerPoint.Application") 
    # Make ConfirmationTools instance (not visible) 
    ct = ConfirmationTools(PowerPoint.Presentations.Open(conf_path, WithWindow=False)) 

    # Initialize reports,
    # then make a cover slide, copy/paste relevant slides, 
    # and save for each report
    reports = init_reports(ct, sim_dir) 
    for rep in reports:
        make_cover(ct, rep)
        copy_slides(ct, rep)
        save_report(rep)

    ct.pptx.Close() # Close, to avoid file corruption, w/o saving
    PowerPoint.Quit() # Quit PowerPoint process

# __all__ = ["_get_rep_type", "_load_template_paths", "_get_date"]

