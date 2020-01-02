import os
import win32com.client as win32

# from time import sleep
# from abc import ABC, abstractmethod
from util import get_interfaces
from reports import ConfirmationTools
from reports.sim import SIReport, PIReport, EMCReport



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
        # "thermal": "",
    }

    # Fetch paths from text file in same directory
    with open(file_path, "r") as f:
        lines = [ line.strip() for line in f.readlines() ]
        for rep_type in templates.keys():
            # Iterate over each line prefixed with template key and = (e.g. "si=")
            for line in lines:
                if line.startswith(rep_type):
                    start = line.index("=")
                    # Omit delimiter and copy remaining str into dict
                    templates[rep_type] = line[start+1:] 

    print("\nLoaded the following templates:\n")
    for k, v in templates.items():
        print(f"  REPORT TYPE {k.upper()}: {v}") 
    print()
    return templates


def init_reports(PowerPoint, conf_tools, sim_dir=""):
    """
    Initializes and returns Report based on user input and template
    """
    templates = _load_template_paths(os.getenv("TEMP_PATH"))
    reports = None
    proj_num = conf_tools.proj_num[:]
    rep_type = conf_tools.type
    template_pptx = PowerPoint.Presentations.Open(templates[rep_type])

    # Instantiate report based on user input
    if rep_type == "si":
        reports = [ SIReport(template_pptx, interface, proj_num) for interface in get_interfaces(conf_tools, sim_dir) ]
    elif rep_type == "pi":
        reports = PIReport(template_pptx, proj_num)
    elif rep_type == "emc":
        reports = EMCReport(template_pptx, proj_num)

    # Ensure returned object is of consistent data structure
    if not isinstance(reports, list):
        reports = [ reports ]

    return reports


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
    reports = init_reports(PowerPoint, ct, sim_dir) 
    for rep in reports:
        rep.build_pptx(ct)

    ct.pptx.Close() # Close, to avoid file corruption, w/o saving
    PowerPoint.Quit() # Quit PowerPoint process

