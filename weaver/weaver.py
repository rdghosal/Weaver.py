import os, sys
import win32com.client as win32

# from time import sleep
# from abc import ABC, abstractmethod
from datetime import date
from .util import *
from .reports import ConfirmationTools, SimulationReport
from .reports.sim import SIReport, PIReport, EMCReport


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////


def fetch_interfaces(dir_path):
    """
    Crawls simulation directory to fetch names of interfaces
    """
    if not os.path.isabs(dir_path) or \
        os.path.split(dir_path)[1] != "Simulation":
        print("ERROR: Simulation folder could not be found.")
        sys.exit(-1)

    return os.listdir(dir_path)


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


def init_reports(conf_tools):
    """
    Initializes and returns Report based on user input and template
    """
    templates = _load_template_paths(TXT_PATH)
    reports = None
    proj_num = conf_tools.proj_num[:]
    rep_type = conf_tools.type

    # Instantiate report based on user input
    if rep_type == "si":
        reports = [ SIReport(templates[rep_type], interface, proj_num) for interface in conf_tools.get_interfaces() ]
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


def weave_reports(conf_path):
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
    reports = init_reports(ct) 
    for rep in reports:
        make_cover(ct, rep)
        copy_slides(ct, rep)
        save_report(rep)

    ct.pptx.Close() # Close, to avoid file corruption, w/o saving
    PowerPoint.Quit() # Quit PowerPoint process

# __all__ = ["_get_rep_type", "_load_template_paths", "_get_date"]

