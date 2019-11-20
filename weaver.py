#!usr/bin/env python
import os, sys, argparse, re

from abc import ABC, abstractmethod
from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** GLOBAL CONSTANTS ****
# //////////////////////////////

# For cover slide
COVER_SLIDE = 0
TITLE_SHAPE = 0 # Same for other slides
DATE_SHAPE = 1 
CREATORS_TABLE = 2
TABLE_COORDS = {
    "preparers": (0,1),
    "reviewers": (1,1),
    "approvers": (2,1) 
}

# slide index of table of contents
TOC = 2


# \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** FUNCTION DEFINITIONS ****
# //////////////////////////////

def load_template_paths(file_path):
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
                    templates[rep_type] = line[start+1:] # Omit delimiter and copy remaining str into dict
    
    return templates
            

def fetch_interfaces(dir_path):
    """
    Crawls simulation directory to fetch names of interfaces
    """
    if not os.path.isabs(dir_path) or \
        os.path.split(dir_path)[1] != "Simulation":
        print("ERROR: Simulation folder could not be found.")
        sys.exit(-1)

    return os.listdir(dir_path)


def get_report_type():
    """
    Returns a valid report type from user
    """
    sim_types = ["si", "pi", "emc", "thermal"] # Types of PCB simulation reports
    while True:
        rep_type = input("Input type of report (SI / PI / EMC / Thermal): " )
        # Verify user input
        if rep_type.lower() in sim_types:
            break

    return rep_type


def get_creators(conf_tools):
    """
    Gets list of authors, reviewers, and approvers from Confirmation Tools object
    """
    creators = {
        "preparers": "",
        "reviewers": "",
        "approvers": ""
    }

    for coords in TABLE_COORDS.values():
        for party in creators:
            creators[party] = conf_tools.pptx.slides[COVER_SLIDE].\
                              shapes[CREATORS_TABLE].table.cell(coords).text

    return creators


def get_date():
    """
    Returns a user inputted date formatted into Japanese
    """
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
    """
    Returns user input title for a particular file
    """
    title = ""
    while True:
       title = input(f"Input title of {target_file}: ")
       if title:
           break

    return title


def init_report(interface, templates, conf_tools):
    """
    Initializes and returns Report based on user input and template
    """
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
    """
    Gets filename from user and closes report after saving
    """
    filename = ""
    while True:
        filename = input("Input filename to save the report: ")
        if filename:
            break

    report.pptx.save(filename)
    print(f"{filename} saved in {path}.")


def make_cover(report):
    """
    Generates first slide of report
    """
    # Used to map creators to correct cell in table

    cover = report.pptx.slides[COVER_SLIDE]
    cover.shapes[TITLE_SHAPE].text = report.title
    cover.shapes[DATE_SHAPE].text = report.date

    # Match table coordinates with Report.creator keys and insert values of latter
    for key, coords in TABLE_COORDS.items():
        for party in report.creators: 
            if key == party:
                cover.shapes[CREATORS_TABLE].cell(coords).text = report.creators[party]

    print(f"Cover slide generated for {report.title}.")


def copy_slides(conf_tools, report):
    """
    Copies over slides of interest 
    from confirmation tools to report
    """
    toc = conf_tools.table_of_contents.copy() # Avoid side effects
    # Read each slide according to toc
    for section in toc:
        for slide_nums in toc[section]:
            for slide_num in slide_nums:
                # Get curr slide of conf_tools
                conf_slide = conf_tools.pptx.slides[slide_num]
                
                # Add a slide of the same format in the report
                layout_name = conf_tools.pptx.slides[slide_num].slide_layout.name
                layout = report.pptx.slides.slide_layout.get_by_name(layout_name)
                report.pptx.slides.add_slide(layout)

                # Fetch newly appended slide
                new_index = len(report.pptx.slides) - 1
                rep_slide = report.pptx.slides[new_index]

                # Copy title of conf_tools slide to report slide
                rep_slide.shapes[TITLE_SHAPE].text = conf_slide.shapes[TITLE_SHAPE].text

                __copy_shapes(conf_slide, conf_tools.slides.index(conf_slide), rep_slide)

    return report


def __copy_shapes(src_slide, src_index, dest_slide):
    """
    Iterates over source slide's shapes 
    and creates matching shapes in destination slide
    Returns None (mutates dest_slide)
    """
    curr = 1 # Starts at 1 to exclude title shape
    for shape in src_slide.shapes[TITLE_SHAPE+1:]:
        # Read position of original
        top = shape.top
        left = shape.left
        height = shape.height
        width = shape.width                    

        if shape.has_text_frame:
            # Add new textbox and add text thereto
            dest_slide.shapes.add_textbox(left, top, width, height)
            dest_slide.shapes[curr].text = shape.text
            curr += 1

        elif shape.has_table:
            # Extract data from table
            table = shape.table
            cols = len(table.columns)
            rows = len(table.rows)
            cells = table.iter_cells() # Cell generator w/ text prop

            # Make table
            dest_slide.shapes.add_table(rows, cols, left, top, width, height)

            # Copy over contents from original to new table
            for cell in cells:
                for new_cell in dest_slide.shapes[curr].table.iter_cells():
                    new_cell.text = cell.text
            curr += 1

        else:
            # Likely a group shape
            try:
                dest_slide.shapes.add_group_shape()
                __copy_group_shape(shape, dest_slide.shapes[curr])
                curr += 1

            except AttributeError: # In case not group shape
                print(f"Unidentifiable shape found in slide {src_index}")


def __copy_group_shape(src_shape, dest_shape):
    """
    Iterates over source group shape
    and creates matching shapes in destination group shape
    Returns None (mutates dest_shape)
    """
    curr_sub = 0 # To track subshapes added to group shape
    img_cnt = 1 # To differentiate img filenames
    for shape in src_shape.shapes:
        # Get position and dimensions of subshape
        sub_left = shape.left
        sub_top = shape.top
        sub_width = shape.width
        sub_height = shape.height

        if shape.shape_type == 13: # PICTURE
            # Make temp folder to extract and save imgs as files
            blob = shape.image.blob
            if not "temp" in os.listdir("."):
                os.mkdir("temp")
            filename = f"temp_img-{img_cnt}.png"
            img_path = os.path.join(os.getcwd(), "temp", filename)
            with open(img_path, "wb") as f:
                f.write(blob)
            dest_shape.shapes.add_picture(filename, sub_left, sub_top,\
                                          sub_width, sub_height)
            img_cnt += 1

        elif shape.shape_type == 17: # TEXT_BOX
            dest_shape.shapes.add_textbox(sub_left, sub_top, sub_width, sub_height)
            dest_shape.shapes[curr_sub].text = shape.text
        
        elif shape.shape_type == None: 
            # Assuming shape is Connector
            begin = shape.begin_x, shape.begin_y
            end = shape.end_x, shape.end_y
            # Default to Straight
            dest_shape.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, begin, end)
            dest_shape

        curr_sub += 1

    # Cleanup
    if "temp" in os.listdir("."):
        os.rmdir("temp")


def weave_reports(templates, interfaces):
    pass


# def make_TOU(report):
#     """Generates Table of Updates"""
#     pass


# def make_TOC(report):
#     """Generates Table of Contents"""
#     pass


# \\\\\\\\\\\\\\\\\\\\\\\\\\\
# *** CLASS DEFINITIONS ****
# ///////////////////////////

# ===============
# Report classes
# ===============

class Report(ABC):
    """
    Base class for simulation report
    """
    def __init__(self, title, template):
        self.__pptx = Presentation(template)
        self.__title = title
        self.__creators = {}
        self.__date = ""
    
    @property
    def pptx(self):
        return self.__pptx

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
    """
    Class for initial, pre-simulation report
    """
    def __init__(self, title, file):
        super().__init__(title, "conf_tools", file)
        self.__creators = get_creators(self) # Override base class

    def fetch_toc(self):
        # Memoize slide num for slides of interest
        toc = self.pptx.slides[TOC].shapes[0].table # Save table
        # To be populated with slide nums
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
                    slide_num = toc.cell(1, y).text
                    toc_dict["sim_targets"] = slide_num
                elif section_name.find("mask") > -1:
                    slide_num = toc.cell(1, y).text
                    toc_dict["eye_masks"] = slide_num
                elif section_name.find("topology") > -1:
                    slide_num = toc.cell(1, y).text
                    toc_dict["topology"] = slide_num
                # Move down TOC
                y += 1 
            # Check if end of TOC in order to end loop
            elif section_name == "" or IndexError:
                break

        # Convert str slide_nums to int for slide indexing
        for slide_num in toc_dict.values():
            if slide_num.find("-") > -1:
                slide_num.split("-")
                slide_num = [ int(num) - 1 for num in slide_num ]
            else:
                slide_num = list(int(slide_num) - 1) # Keep data structure consistent
        
        return toc_dict

    # def list_interfaces(self):
    #     interfaces = []
    #     toc = self.__toc.copy() # To avoid side effects
    #     start = 0
    #     end = 1

    #     # Start and end indicies matched to self.table_of_contents
    #     if len(toc["sim_targets"]) > 1:
    #         start = toc["sim_targets"][0]
    #         end += toc["sim_targets"][1]
    #     else:
    #         start = toc["sim_targets"]
    #         end += start

    #     for slide in self.pptx.slides[start:end]:
    #         try:
    #             title = slide.shapes[TITLE_SHAPE].lower()
    #             match = re.search(r"condition\s?:\s?(\w+)\b", title)
    #             # Interfaces are found in section 2 only
    #             if match:
    #                 interfaces.append(match.group(1)) # Save capture group (viz. interface)
    #         except:
    #             continue

    #     # Remove duplicates
    #     return set(interfaces)


class SimulationReport(Report):
    """
    Base class for simulation reports
    """
    def __init__(self, title, template, sim_type):
        super().__init__(title, template)
        self.__sim_type = sim_type

    @property
    def simulation_type(self):
        return self.__sim_type


class SIReport(SimulationReport):
    """
    Class for PCB signal integrity report
    """
    def __init__(self, title, template):
        super().__init__(title, "SI", template)


class PIReport(SimulationReport):
    """
    Class for PCB power integrity report
    """
    def __init__(self, title, template):
        super().__init__(title, "PI", template)
        self.__net_names = []

    @property
    def net_names(self):
        return self.__net_names

    @net_names.setter
    def net_names(self, value):
        self.__net_names = value


class EMCReport(SimulationReport):
    """
    Class for PCB EMC report
    """
    def __init__(self, title, template):
        super().__init__(title, "EMC", template)


class ThermalReport(SimulationReport):
    """
    Class for PCB thermal report
    """
    def __init__(self, title, template):
        super().__init__(title, "Thermal", template)


# # ==============
# # Slide classes
# # ==============

# class Slide():
#     """Base class for slide in report"""
#     pass


# class CoverSlide(Slide):
#     """Class for first slide of report"""
#     pass


# class DividerSlide(Slide):
#     """Class for slide dividing sections of report"""
#     pass



# # ========================
# # SlideContent Base Class
# # ========================

# class SlideContent():
#     """Base class for images, textboxes, tables, etc. on slides"""
#     def __init__(self, wh, hasBorder, xy):
#         self.__dimensions = wh # tuple
#         self.__hasBorder = hasBorder
#         self.__position = xy # tuple
#         self.__border = {
#             "type": "",
#             "thickness": 0,
#             "style": "",
#             "color": ""
#         }
    
#     def set_border(self):
#         raise NotImplementedError


# # ==============
# # Image classes
# # ==============

# class Image(SlideContent):
#     """Base class for images on report"""
#     pass

# class TopologyImage(Image):
#     """Topology diagram"""
#     pass


# class WaveForm(Image):
#     """Waveform images for SI reports"""
#     pass


# # ==============
# # TextBox classes
# # ==============

# class TextBox(SlideContent):
#     """Base class for labels on slides"""
#     def __init__(self, ff, color, bg_color, size, hasBorder, position):
#         super().__init__(size, hasBorder, position)
#         self.__font_family = ff
#         self.__font_color = color
#         self.__bg_color = bg_color
    
    
# class Title(TextBox):
#     pass 


# class Subtitle(Title):
#     pass


# class Label(TextBox):
#     """All labels besides (sub)titles"""
#     def __init__(self, color):
#         super().__init__(color)

# class Comment(TextBox):
#    pass 


# # ==============
# # TextBox classes
# # ==============

# class Table(SlideContent):
#     """Base class for tables on report slides"""
#     def __init__(self, w, h):
#         self.__num_rows = h
#         self.__num_cols = w


# class Subtable(Table):
#     def __init__(self, w, h):
#         super().__init__(w, h)    


# class CellCollection():
#     def __init__(self, colors):
#         self.__cell_color = colors[0]
#         self.__font_color = colors[1]

# class Column(CellCollection):
#     def __init__(self, colors):
#         super().__init__(colors)


# class Row(CellCollection):
#     def __init__(self, colors):
#         super().__init__(colors)


# class Header(Row):
#     pass


if __name__ == "__main__":
    desc = """
            Weaver.py takes paths to: 
                (1) a confirmation tools report,
                (2) a Simulaton directory
                (3) a templates.txt listing templates to be used (optional)
                (4) a destination directory (optional)

            to automatically generate simulation reports 
            according to a templates listed in (3).

            For more information refer to the README.
           """
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument("simulation_directory", help="Simulation directory") 

    args = parser.parse_args()   

    templates = load_template_paths(args.textfile)
    interfaces = fetch_interfaces(args.simulation_directory)

    weave_reports(templates, interfaces)