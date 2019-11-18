#!usr/bin/env python
import os, argparse

from abc import ABC, abstractmethod
from pptx import Presentation


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////

def get_report_type():
    """Returns a valid report type from user"""
    sim_types = ["si", "pi", "emc", "thermal"] # Types of PCB simulation reports
    while True:
        r_type = input("Input type of report (SI/ PI / EMC / Thermal): " )
        # Verify user input
        if r_type.lower() in sim_types:
            break

    return r_type


def get_creators():
    """Gets list of authors, reviewers, and approvers of report"""
    creators = {
        "preparers": [],
        "reviewers": [],
        "approvers": []
    }

    # Ask user to input responsible persons per party
    for party in creators:
        persons = input(f"Input {party} with a comma-separated list:\n").split(",") # Convert user input into list
        creators[party] = [ person.strip() for person in persons ] # Remove extraneous spacing

    return creators


def get_date():
    """Returns a user inputted date formatted into Japanese"""
    date = ""
    while True:
        prompt = """
        Input date as follows:

        yyyy,MM,dd

        where:
        yyyy -> year
        MM -> month
        dd -> date
        """
        try:
            date = input(prompt).split(",")
            break
        except: 
            continue

    return u"{0}年　{1}月　{2}日".format(date[0], date[1], date[2])
        

def get_title():
    """Returns user input title for report"""
    title = ""
    while True:
       title = input("Input title of report: ")
       if title:
           break

    return title


def init_report(template):
    """Initializes and returns Report based on user input and template"""
    # Get basic properties of report
    title = get_title()
    rep_type = get_report_type()

    # Maps to path for report template with rep_type as key
    templates = {
        "si": "path1",
        "pi": "path2",
        "emc": "path3",
        "thermal": "path4"
    }

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

    report.creators = get_creators()         
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
    table_coords = {
        "preparers": (0,1),
        "reviewers": (1,1),
        "approvers": (2,1) 
    }

    cover = report.pptx.slides[0]
    cover.shapes[0].text = report.title
    cover.shapes[1].text = report.date

    # Match table coordinates with Report.creator keys and insert values of latter
    for key, coords in table_coords.items():
        for party in report.creators: 
            if key == party:
                cover.shapes[2].cell(coords).text = report.creators[party]

    print(f"Cover slide generated for {report.title}")


def make_TOU(report):
    """Generates Table of Updates"""
    pass


def make_TOC(report):
    """Generates Table of Contents"""
    pass


# \\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** CLASS DEFINITIONS ****
# ///////////////////////////

# ========================
# ConfirmationTools class
# ========================
class ConfirmationTools():
    def __init__(self, path):
        self.__file = Presentation(path)
    
    def get_interfaces(self):
        pass


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
    def creators(self, creators):
        # Format each value of input dict into str
        for party, persons in creators.items():
            persons_str = ""
            for i in range(len(persons)):
                # Add comma unless last item
                delim = ", " if i > len(persons) - 1 else "" 
                persons_str += persons[i] + delim
            creators[party] = persons_str # Replace value with formatted str
        self.__creators = creators

    @property
    def date(self):
        return self.__date

    @date.setter
    def date(self, value):
        self.__date = value


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