import re
from time import sleep
from ..simreport import SimulationReport
from util import MSOTRUE

SIM_TARGET = 6
SIM_TARGET_REP = 7
DC_DROP = 12
AC_DROP = 13
IMPEDANCE = 14

class PIReport(SimulationReport):
    """
    Class for PCB power integrity report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)
        self.__power_nets = {}
        self.__counter = 1

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
    
    def _copy_slides(self, conf_tools):
        toc = conf_tools.get_toc()
        pages = ( toc["sim_target"][0], toc["voltage_margin"][1] )
        for i in range(pages[0], pages[1] + 1):
            # Skip impedance table
            if i - pages[0] == 1:
                i += 1
            conf_tools.pptx.Slides(i).Copy()
            sleep(.25)
            self.pptx.Slides.Paste(i + 1) # Offset by one
            self.__counter += 1

    def _read_power_nets(self):
        slide = self.pptx.Slides(SIM_TARGET_REP)
        table = self._get_table(slide.Shapes)

        for i in range(2, len(table.Rows)):
            net = {
                "power net": "",
                "reference ic": "",
                "voltage": "",
                "dc drop analysis": False,
                "ac drop analysis": [ False, "" ],
                "impedance analysis": [ False, "" ],
                "acceptable target voltage margin": ""
            }

            for j in range(1, len(table.Columns) + 1):
                col_name = table.Cell(1, j).Shape.TextFrame.TextRange.Text[:].lower()
                if col_name.find("\r") > -1:
                    col_name = " ".join(col_name.split("\r"))
                text = table.Cell(i, j).Shape.TextFrame.TextRange.Text[:]
                if col_name in ["ac drop analysis", "impedance analysis"]:
                    match = re.search(r"○\s*\((.+)\)", text)
                    if match:
                        net[col_name][0] = True
                        load = match.group(1) # Get load IC
                        net[col_name][1] = load 
                elif col_name == "dc drop analysis" and text.find("○") > -1:
                    net[col_name] = True
                else:
                    net[col_name] = text[:]

            print("\nLoaded the following net:\n")
            for k, v in net.items():
                print(f"  {k.upper()}: {v}\n")
            print("---------------" * 5)
            yield net
    
    def _parse_net_info(self, net, analysis_type, item_num):
        tar_info = net[analysis_type] if analysis_type == "dc drop analysis" else net[analysis_type][0]
        if not tar_info: 
            return None
        reference = net["reference ic"][:]
        # Change in case load is set to "all"
        if analysis_type.startswith("ac"):
            if net["reference ic"].lower().find("all load ic"):
                reference = reference.split("~")[0]
                reference += net[analysis_type][1][:]
        elif analysis_type.startswith("imp"):
            reference = reference.split("~")[0]
            reference += net[analysis_type][1][:]
        # Info to be filled into table
        net_info = {
            "no.": item_num,
            "power net": net["power net"],
            "reference ic": reference,
            "source voltage": net["voltage"]
        }
        return net_info

    def _fill_analysis_tables(self, type_):

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
        for n in self._read_power_nets():
            target_nets.append(self._parse_net_info(n, anal_type, item_num))
            item_num += 1
        
        slide = self.pptx.Slides(index)
        table = self._get_table(slide.Shapes)

        while len(table.Rows) < len(target_nets):
            table.Rows.Add()
        
        num_cols = 3 if type_ == "imp" else 4
        for i in range(len(target_nets)):
            # Only iterating first four columns
            if not target_nets[i]:
                continue
            for j in range(num_cols):
                col = j + 1
                row = i + 3
                header = 1 if col <= 1 else 2 # To accomodate different header sizes
                col_name = table.Cell(header, col).Shape.TextFrame.TextRange.Text[:].lower()
                # To ensure consistency with net_info dict
                # if analysis type is impedance
                if not type_ == "imp":
                    if col_name == "simulation target":
                        col_name = "power net"
                    elif col_name == "simulation portion":
                        col_name = "reference ic"
                    elif col_name == "item":
                        col_name = "no."
                try:
                    table.Cell(row, col).Shape.TextFrame.TextRange.Text = target_nets[i][col_name]
                except KeyError:
                    break
    
    def _replace_placeholders(self, net, shape):
        placeholders = {
            "<POWER_NET[i]>": net["power net"],
            "<V[i]>": net["voltage"],
            "<RECEIVER_REF>": net["reference ic"]
        }
        # Look at TextFrame of shape or cell of table therewithin
        if shape.HasTextFrame == MSOTRUE:
            curr_text = shape.TextFrame.TextRange.Text[:]
            for k in placeholders.keys():
                if curr_text.find(k) > -1:
                    curr_text = curr_text.replace(k, placeholders[k])
            shape.TextFrame.TextRange.Text = curr_text
        elif shape.HasTable == MSOTRUE:
            tar_cells = [(2,1), (3,1)]
            for cell in tar_cells:
                curr_text = shape.Table.Cell(cell[0], cell[1]).Shape.TextFrame.TextRange.Text[:]
                for k in placeholders.keys():
                    if curr_text.find(k) > -1:
                        shape.Table.Cell(cell[0], cell[1]).Shape.TextFrame.TextRange.Text =\
                            curr_text.replace(k, placeholders[k])

    def _build_slides(self):
        # Set ptrs to three result type slides
        slide_ptrs = {
            "dc drop analysis": 13 + self.__counter,
            "ac drop analysis": 14 + self.__counter,
            "impedance analysis": 15 + self.__counter
        }

        self._curr_slide = slide_ptrs["impedance analysis"] + 1
        for net in self._read_power_nets():
            for analysis in ["dc drop analysis", "ac drop analysis", "impedance analysis"]:
                target = net[analysis] if analysis == "dc drop analysis" else net[analysis][0]
                if target: 
                    self.pptx.Slides(slide_ptrs[analysis]).Copy()
                    sleep(.25)
                    self.pptx.Slides.Paste(self._curr_slide)
                    for shape in self.pptx.Slides(self._curr_slide).Shapes:
                        self._replace_placeholders(net, shape)
                    self._curr_slide += 1
        
        # Remove template slide
        for v in slide_ptrs.values():
            self.pptx.Slides(v).Delete()

    def build_pptx(self, conf_tools):
        self._make_cover(conf_tools)
        self._copy_slides(conf_tools)
        for type_ in ["ac", "dc", "imp"]: self._fill_analysis_tables(type_)
        self._build_slides()
        self._save_report()




    
