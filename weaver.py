#!usr/bin/env python
import os, argparse, re

from abc import ABC, abstractmethod
from pptx import Presentation


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** GLOBAL CONSTANTS ****
# //////////////////////////////

# For cover page
COVER_PAGE = 0
TITLE_SHAPE = 0 # Same for other pages
DATE_SHAPE = 1 
CREATORS_TABLE = 2
TABLE_COORDS = {
    "preparers": (0,1),
    "reviewers": (1,1),
    "approvers": (2,1) 
}

# Page index of table of contents
TOC = 2


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////

def load_template_paths():
    """Fetches template paths and returns dict mapping report type to template"""
    # dict to be populated
    templates = {
        "si": "",
        "pi": "",
        "emc": "",
        "thermal": "",
    }

    # Fetch paths from text file in same directory
    with open("template_paths.txt") as f:
        for rep_type in templates:
            # Iterate over each line prefixed with template key and = (e.g. "si=")
            for line in f.readlines():
                if line.startswith(rep_type):
                    start = line.index("=")
                    templates[rep_type] = line[start+1:] # Omit delimiter and copy remaining str into dict
    
    return templates
            


def get_report_type():
    """Returns a valid report type from user"""
    sim_types = ["si", "pi", "emc", "thermal"] # Types of PCB simulation reports
    while True:
        rep_type = input("Input type of report (SI / PI / EMC / Thermal): " )
        # Verify user input
        if rep_type.lower() in sim_types:
            break

    return rep_type


def get_creators(conf_tools):
    """Gets list of authors, reviewers, and approvers from Confirmation Tools object"""
    creators = {
        "preparers": "",
        "reviewers": "",
        "approvers": ""
    }

    for coords in TABLE_COORDS.values():
        for party in creators:
            creators[party] = conf_tools.pptx.slides[COVER_PAGE].\
                              shapes[CREATORS_TABLE].table.cell(coords).text

    return creators


def get_date():
    """Returns a user inputted date formatted into Japanese"""
    date = ""
    while True:
        # Instructions for user input
        prompt = """
        Input report date as follows:

        yyyy,MM,dd

        where:
        yyyy -> year
        MM -> month
        dd -> date
        """
        # Check if instructions were followed
        try:
            date = input(prompt).split(",") # Split for output str formatting
            break
        except: 
            continue

    return u"{0}年　{1}月　{2}日".format(date[0], date[1], date[2])
        

def get_title(target_file):
    """Returns user input title for a particular file"""
    title = ""
    while True:
       title = input(f"Input title of {target_file}: ")
       if title:
           break

    return title


def make_rep_structure(conf_tools):
    """Opens conf_tools and """
    pass


# Made global to avoid unnecessary function call
templates = load_template_paths()

def init_report(interface, templates, conf_tools):
    """Initializes and returns Report based on user input and template"""
    # Get basic properties of report
    title = get_title(f"simulation report for the interface {interface}")
    rep_type = get_report_type()

    report = None
    # Instantiate report based on user input
    if rep_type == "si":
        report = SIReport(title, templates[rep_type])
    elif rep_type == "pi":
        report = PIReport(title, templates[rep_type])
    elif rep_type == "emc":
        report = EMCReport(title, templates[rep_type])
    else:
        report = ThermalReport(title, templates[rep_type])

    report.creators = conf_tools.creators 
    report.date = get_date()

    return report


def save_report(report, path):
    """Gets filename from user and closes report after saving"""
    filename = ""
    while True:
        filename = input("Input filename to save the report: ")
        if filename:
            break
    report.file.save(filename)
    print(f"{filename} saved in {path}.")


def make_cover(report):
    """Generates first page of report"""
    # Used to map creators to correct cell in table

    cover = report.pptx.slides[COVER_PAGE]
    cover.shapes[TITLE_SHAPE].text = report.title
    cover.shapes[DATE_SHAPE].text = report.date

    # Match table coordinates with Report.creator keys and insert values of latter
    for key, coords in TABLE_COORDS.items():
        for party in report.creators: 
            if key == party:
                cover.shapes[CREATORS_TABLE].cell(coords).text = report.creators[party]

    print(f"Cover slide generated for {report.title}.")


def clone_pages(conf_tools, report):
    # 1. name = c_t.slides[i].slide_layout.name
    # 2. layout = pptx.slides.slide_layout.get_by_name(name) 
    # 3. rep.pptx.slides.add_slide(layout)
    # 4. Transfer contents of matching shapes
    # >> If not in layout (e.g. pictures):
    # 1. pics (pictures/groupshape), table(graphicframe) -> get top, left of orig (w, h?)
    # 2. text -> top, left, and shape.text_frame.paragraph.font (None, okay?)
   
    toc = conf_tools.table_of_contents.copy() # Avoid side effects

    # Read each page according to toc
    for section in toc:
        for pg_nums in toc[section]:
            for pg_num in pg_nums:
                # Get collection of all shapes in conf_tools
                ct_slide = conf_tools.ppt.slides[pg_num]
                
                # Add a slide of the same format in the report
                layout_name = conf_tools.pptx.slides[pg_num].slide_layout.name
                layout = report.pptx.slides.slide_layout.get_by_name(layout_name)
                report.pptx.slides.add_slide(layout)

                # Fetch slide added to report
                new_index = len(report.pptx.slides) - 1
                rep_slide = report.pptx.slides[new_index]

                # Copy title of conf_tools slide to report slide
                rep_slide.shapes[TITLE_SHAPE].text = ct_slide.shapes[TITLE_SHAPE].text
                


def make_TOU(report):
    """Generates Table of Updates"""
    pass


def make_TOC(report):
    """Generates Table of Contents"""
    pass


# \\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** CLASS DEFINITIONS ****
# ///////////////////////////

# ===============
# Report classes
# ===============

class Report(ABC):
    """Base class for simulation report"""
    def __init__(self, title, rep_type, template):
        self.__pptx = Presentation(template)
        self.__type = rep_type
        self.__title = title
        self.__creators = {}
        self.__date = ""
    
    @property
    def pptx(self):
        return self.__pptx

    @property
    def report_type(self):
        return self.__type
    
    @property
    def title(self):
        return self.__title

    @property
    def creators(self):
        return self.__creators

    @creators.setter
    def creators(self, value):
        self.__creators = value

    @property
    def date(self):
        return self.__date

    @date.setter
    def date(self, value):
        self.__date = value


class ConfirmationTools(Report):
    """Class for initial, pre-simulation report"""
    def __init__(self, title, file):
        super().__init__(title, "conf_tools", file)
        self.__creators = get_creators(self) # Override base class
        self.__toc = {}

    @property
    def table_of_contents(self):
        return self.__toc

    def read_toc(self):
        # Memoize page num for pages of interest
        toc = self.pptx.slides[TOC].shapes[0].table # Save table
        # To be populated with page nums
        toc_dict = {
            "sim_targets": None,
            "eye_masks": None,
            "topology": None
        }

        y = 1 # Starting y coord of table traversal
        while True:
            section_name = toc.cell(0, y).text.lower()
            # Only contents in section 2 is of interest
            if section_name.startswith("2"):
                if section_name.find("target") > -1:
                    page_num = toc.cell(1, y).text
                    toc_dict["sim_targets"] = page_num
                elif section_name.find("mask") > -1:
                    page_num = toc.cell(1, y).text
                    toc_dict["eye_masks"] = page_num
                elif section_name.find("topology") > -1:
                    page_num = toc.cell(1, y).text
                    toc_dict["topology"] = page_num
                # Move down TOC
                y += 1 
            # Check if end of TOC in order to end loop
            elif section_name == "" or IndexError:
                break

        # Convert str page_nums to int for slide indexing
        for page_num in toc_dict.values():
            if page_num.find("-") > -1:
                page_num.split("-")
                page_num = [ int(num) - 1 for num in page_num ]
            else:
                page_num = list(int(page_num) - 1) # Keep data structure consistent
        
        self.__toc = toc_dict

    def list_interfaces(self):
        interfaces = []
        toc = self.__toc.copy() # To avoid side effects
        start = 0
        end = 1

        # Start and end indicies matched to the self.table_of_contents
        if len(toc["sim_targets"]) > 1:
            start = toc["sim_targets"][0]
            end += toc["sim_targets"][1]
        else:
            start = toc["sim_targets"]
            end += start

        for slide in self.pptx.slides[start:end]:
            try:
                title = slide.shapes[TITLE_SHAPE].lower()
                match = re.search(r"condition\s?:\s?(\w+)\b", title)
                # Interfaces are found in section 2 only
                if match:
                    interfaces.append(match.group(1)) # Save capture group (viz. interface)
            except:
                continue

        # Remove duplicates
        return set(interfaces)
    
    def fetch_topologies(self, interface):
        pass


class SIReport(Report):
    """Class for PCB signal integrity report"""
    def __init__(self, title, template):
        super().__init__(title, "SI", template)


class PIReport(Report):
    """Class for PCB power integrity report"""
    def __init__(self, title, template):
        super().__init__(title, "PI", template)
        self.__net_names = []

    @property
    def net_names(self):
        return self.__net_names

    @net_names.setter
    def net_names(self, value):
        self.__net_names = value


class EMCReport(Report):
    """Class for PCB EMC report"""
    def __init__(self, title, template):
        super().__init__(title, "EMC", template)


class ThermalReport(Report):
    """Class for PCB thermal report"""
    def __init__(self, title, template):
        super().__init__(title, "Thermal", template)


# ==============
# Slide classes
# ==============

class Slide():
    """Base class for slide in report"""
    pass


class CoverSlide(Slide):
    """Class for first slide of report"""
    pass


class DividerSlide(Slide):
    """Class for slide dividing sections of report"""
    pass



# ========================
# SlideContent Base Class
# ========================

class SlideContent():
    """Base class for images, textboxes, tables, etc. on slides"""
    def __init__(self, wh, hasBorder, xy):
        self.__dimensions = wh # tuple
        self.__hasBorder = hasBorder
        self.__position = xy # tuple
        self.__border = {
            "type": "",
            "thickness": 0,
            "style": "",
            "color": ""
        }
    
    def set_border(self):
        raise NotImplementedError


# ==============
# Image classes
# ==============

class Image(SlideContent):
    """Base class for images on report"""
    pass

class TopologyImage(Image):
    """Topology diagram"""
    pass


class WaveForm(Image):
    """Waveform images for SI reports"""
    pass


# ==============
# TextBox classes
# ==============

class TextBox(SlideContent):
    """Base class for labels on slides"""
    def __init__(self, ff, color, bg_color, size, hasBorder, position):
        super().__init__(size, hasBorder, position)
        self.__font_family = ff
        self.__font_color = color
        self.__bg_color = bg_color
    
    
class Title(TextBox):
    pass 


class Subtitle(Title):
    pass


class Label(TextBox):
    """All labels besides (sub)titles"""
    def __init__(self, color):
        super().__init__(color)

class Comment(TextBox):
   pass 


# ==============
# TextBox classes
# ==============

class Table(SlideContent):
    """Base class for tables on report slides"""
    def __init__(self, w, h):
        self.__num_rows = h
        self.__num_cols = w


class Subtable(Table):
    def __init__(self, w, h):
        super().__init__(w, h)    


class CellCollection():
    def __init__(self, colors):
        self.__cell_color = colors[0]
        self.__font_color = colors[1]

class Column(CellCollection):
    def __init__(self, colors):
        super().__init__(colors)


class Row(CellCollection):
    def __init__(self, colors):
        super().__init__(colors)


class Header(Row):
    pass