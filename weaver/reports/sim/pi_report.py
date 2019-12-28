from ..simreport import SimulationReport
import re

SIM_TARGET = 6
SIM_TARGET_REP = 7
DC_DROP = 13
AC_DROP = 14
IMPEDANCE = 15

class PIReport(SimulationReport):
    """
    Class for PCB power integrity report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)
        self.__power_nets = {}

    def __str__(self):
        pass

    @property
    def title(self):
        return f"{self.proj_num}\nVerification of Power Integrity by PI Simulation [Ver.1.0]"

    @property
    def report_type(self):
        return "PI"
    
    @property
    def net_names(self):
        return self.__power_nets[:]

    def _get_power_nets(self):
        # TODO reuse from EMC and abstract to SimulationReport
        pass
    
    def _get_duplicable_nums(self, conf_tools, first_sec, last_sec):
        table = self._get_table(conf_tools.Slides(3).Shapes)
        start, end = 0, 0
        row = 4 # init row
        while True:
            text = table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]
            if text.find(first_sec) > -1:
                start = table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                # Take first value
                if start.find("-"):
                    start = start.split("-")[0]
                elif start.find("\u2013"):
                    start = start.split("\u2013")[0]
                start = int(start)
            elif text.find(last_sec) > -1:
                end = table.Cell(row, 2).Shape.TextFrame.TextRange.Text[:]
                # Take last value
                if end.find("-"):
                    end = end.split("-")[1]
                elif end.find("\u2013"):
                    end = end.split("\u2013")[1]
                end = int(end)
    
    def copy_from_conf(self, pages, conf_tools):
        for i in range(pages[0], pages[1] + 1):
            # Skip impedance table
            if i - pages[0] == 1:
                i += 1
            conf_tools.pptx.Slides(i).Copy()
            self.pptx.Slides.Paste(i + 1) # Offset by one

    def _read_power_nets(self):
        power_nets = []
        slide = self.pptx.Slides(SIM_TARGET_REP)
        table = self._get_table(slide.shapes)

        for i in range(2, len(table.Rows)):
            net = {
                "power net": "",
                "reference ic": "",
                "voltage": "",
                "dc drop analysis": False,
                "ac drop analysis": (False, ""),
                "impedance analysis": (False, ""),
                "acceptable total voltage margin": ""
            }

            for j in range(1, len(table.Columns)):
                col_name = table.Cell(1, j).Shape.TextFrame.TextRange.Text[:].lower()
                text = table.Cell(i, j).Shape.TextFrame.TextRange.Text[:]

                if col_name in ["ac drop anaysis", "impedance analysis"]:
                    match = re.search(r"〇\s*\((\w)\)\b", text)
                    if match:
                        net[col_name][0] = True
                        load = match.group(1) # Get load IC
                        net[col_name][1] = load 
                else:
                    net[col_name] = text[:]

            power_nets.append(net)
    
        self.power_nets = power_nets[:]

    def fill_analysis_table(self, type_):
        if type_ not in ["ac", "dc", "imp"]:
            print("The targeted analysis type is invalid")
            raise AnalysisTypeError

        p_nets = self.power_nets

        index = None # To be used later for finding target slide
        anal_type = ""

        if type_ == "ac":
            index = AC_DROP
            anal_type = "ac drop analysis"
        elif type_ == "dc":
            index = DC_DROP
            anal_type = "dc drop analysis"
        else:
            index = IMPEDANCE
            anal_type = "impedance analysis"

        target_nets = []

        item_num = 1
        for n in p_nets:
            if n[anal_type][0]:
                reference = n["reference ic"][:]
                # Change in case load is set to "all"
                if type_ == "ac":
                    if n["reference ic"].lower().find("all load ic"):
                        reference = reference.split("~")[0]
                        reference += n["ac drop analysis"][1][:]
                elif type_ == "imp":
                    reference = reference.split("~")[0]
                    reference += n["impedance analysis"][1][:]
                # Info to be filled into table
                net_info = {
                    "no.": item_num,
                    "power net": n["power net"],
                    "reference ic": reference,
                    "source voltage": n["voltage"]
                }
                target_nets.append(net_info)
                item_num += 1
        
        slide = self.pptx.Slides(index)
        table = self._get_table(slide.Shapes)

        while len(table.Rows) < len(target_nets):
            table.Rows.Add()
        
        num_cols = 3 if type_ == "imp" else 4
        for i in range(len(target_nets)):
            # Only iterating first four columns
            for j in range(num_cols):
                col = j + 1
                row = i + 2 if col <= 1 else i + 3
                header = 1 if col <= 1 else 2 # To accomodate different header sizes
                col_name = table.Cell(header, col).Shape.TextFrame.TextRange.Text[:].lower()
                # To ensure consistency with net_info dict
                # if analysis type is impedance
                if type_ == "imp":
                    if col_name == "simulation target":
                        col_name = "power net"
                    elif col_name == "simulation portion":
                        col_name = "reference ic"
                    elif col_name == "item":
                        col_name = "no."
                table.Cell(row, col).Shape.TextFrame.TextRange.Text = target_nets[i][col_name][:]
        

class AnalysisTypeError(Exception):
    pass





    
