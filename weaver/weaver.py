import win32com.client as win32

# from time import sleep
# from abc import ABC, abstractmethod
from .util import TXT_PATH, get_interfaces
from .reports import ConfirmationTools
from .reports.sim import SIReport, PIReport, EMCReport



# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////


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


def weave_reports(conf_path, sim_dir):
    """
    Generate reports based on input confirmation tools and indicated type
    """
    # Start PowerPoint process
    PowerPoint = win32.Dispatch("PowerPoint.Application") 
    # Make ConfirmationTools instance (not visible) 
    ct = ConfirmationTools(PowerPoint.Presentations.Open(conf_path, WithWindow=False)) 

    # Initialize reports,
    # then make a cover slide, copy/paste relevant slides, 
    # and save for each report
    reports = init_reports(ct, sim_dir) 
    for rep in reports:
        rep.build_pptx(ct)
        # make_cover(ct, rep)
        # copy_slides(ct, rep)
        # save_report(rep)

    ct.pptx.Close() # Close, to avoid file corruption, w/o saving
    PowerPoint.Quit() # Quit PowerPoint process

