from ..simreport import SimulationReport


class SIReport(SimulationReport):
    """
    Class for PCB signal integrity report
    """
    def __init__(self, template, interface, proj_num):
        super().__init__(template, proj_num)
        self.__interface = interface
    
    def __str__(self):
        return f"{self.report_type} Report for {self.interface}"

    @property
    def title(self):
        return f"{self.proj_num}\nVerification of Signal Integrity\n{self.interface} [Ver.1.0]"

    @property
    def report_type(self):
        return "SI"

    @property
    def interface(self):
        return self.__interface
    
    def _read_interfaces(self, conf_tools):
        toc = conf_tools.get_toc()
        start, end = toc["sim_targets"][0], toc["sim_targets"][1]
        start_slide = self.pptx.Slides(start)
        interfaces = []
        last_title = ""
        for slide in range(start, end + 1):
            for shape in slide.Shapes:
                if shape.HasTextFrame == MSOTRUE:
                    text = shape.TextFrame.TextRange.Text[:].lower()
                    last_title = text[:]
                    if text.find("target") > -1:
                        tar_index = None
                        tar_index = text.index(":")
                        # In case full-size colon used
                        if not tar_index:
                            tar_index = text.find("：")
                        # Displace pointer to the right by 1
                        if_ = text[tar_index+1:]
                        if_ = {
                            "name": "",
                            "signals": [{
                                "name": "",
                                "driver": {
                                    "ref_no": "",
                                    "part_name": "",
                                    "ibis_model": "",
                                    "buffer_model": ""
                                },
                                "receiver": {
                                    "ref_no": "",
                                    "part_name": "",
                                    "ibis_model": "",
                                    "buffer_model": ""
                                },
                                "frequency": "",
                                "PVT": ""
                            }],
                        }
                        # if_ = if_.strip() # clean up
                        interfaces.append(if_)

                        

    