from time import sleep
from ..simreport import SimulationReport
from util import TITLE_NAME, MSOTRUE, com_error

SIM_TARGETS = 6


class EMCReport(SimulationReport):
    """
    Class for PCB EMC report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)
        self.__power_nets = []
        # TODO: implement a toc prop for random access

    def __str__(self):
        pass

    @property
    def title(self):
        return f"{self.proj_num}\nEMC (Power Resonance) Simulation [Ver.1.0]"
    
    @property
    def report_type(self):
        return "EMC"

    @property
    def power_nets(self):
        return self.__power_nets[:]

    def _update_toc(self):
        """Updates table of contents after appending slides to a section"""
        # TODO
    
    def _get_power_nets(self, conf_tools):
        """Reads table of contents (TOC) for pages that need making"""
        toc = conf_tools.get_toc()
        tar_pages = toc["sim_target"]
        # Get page(s) and copy them into report

        new_slides = tar_pages[1] - tar_pages[0]
        self._curr_slide = 6

        for slide in (self._curr_slide, self._curr_slide + new_slides):
            for shape in self.pptx.Slides(slide).Shapes:
                if shape.HasTextFrame == MSOTRUE:
                    sec_num = "3.1"
                    curr_text = shape.TextFrame.TextRange.Text[:]
                    if curr_text.strip().startswith(sec_num):
                        shape.TextFrame.TextRange.Text = curr_text.replace(sec_num, "")
        
        self._curr_slide += new_slides
        
        sim_tar_table = self._get_table(self.pptx.Slides(SIM_TARGETS).Shapes)
        row = 2 # Initial row for scan
        while True:
            try:
                net_col = 1 # For net names
                voltage_col = 3
                resonance_col = 4 # For y/n power resonance analysis
                has_resonance = False 
                net = sim_tar_table.Cell(row, net_col).Shape.TextFrame.TextRange.Text[:]
                voltage = sim_tar_table.Cell(row, voltage_col).Shape.TextFrame.TextRange.Text[:]
                if sim_tar_table.Cell(row, resonance_col).Shape.TextFrame.TextRange.Text[:] == u"〇":
                    has_resonance = True
                self.__power_nets.append((net, voltage, has_resonance))
                row += 1 # Move down table
            # Found end of table
            except com_error:
                break

        return self.power_nets

    def _fill_analysis_table(self):
        """Populates resonance analysis table with power net names"""
        self._curr_slide += 2 # Move past divider (assumes execution after get_power_nets)
        index = self._curr_slide        

        # Grab table from slide
        slide = self.pptx.Slides(index)
        table = self._get_table(slide.Shapes)

        row = 3 # init row
        count = 0
        item_num = 1

        # Make sure table has same number of rows as power nets,
        # excluding the header and accounting for two rows per power net
        while len(table.Rows) - 1 < len(self.power_nets) * 2:
            table.Rows.Add()

        while True:
            try:
                if row % 2 != 0:
                    if row > 3: item_num += 1
                    table.Cell(row, 1).Shape.TextFrame.TextRange.Text = str(item_num)
                new = self.power_nets[item_num - 1][0]
                table.Cell(row, 2).Shape.TextFrame.TextRange.Text = new
                count += 1
                row += 1
            except com_error:
                break

    def _make_reson_analysis(self):
        """Copy template for resonance analysis and fill in table and title"""
        self._curr_slide += 3 # move to next (needs better error-proofing)
        index = self._curr_slide
        shapes = self.pptx.Slides(index).Shapes

        # Exclude init template slide and calc number of times to copy
        num_nets = len(self.power_nets) - 1
        count = 0
        self.pptx.Slides(index).Copy()
        sleep(.25)
        # Use filter to only get those nets that need resonance analysis
        p_nets = self.power_nets
        while count < num_nets:
            self.pptx.Slides.Paste(index + 1 + count) # Place right after current
            shapes = self.pptx.Slides(index + 1 + count).Shapes
            for s in shapes:
                if s.HasTextFrame == MSOTRUE:
                    text = s.TextFrame.TextRange.Text[:]
                    if text.startswith("Target"):
                        # TODO: use boolean to make sure only nets needing resonance analysis are used
                        new = text.replace("<V[i]>", p_nets[count][1][:])
                        new = new.replace("<POWER_NET[i]>", p_nets[count][0][:])
                        s.TextFrame.TextRange.Text = new
                elif s.HasTable == MSOTRUE:
                    s.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = p_nets[count][0]

            # Move to next power net        
            count += 1

        self.pptx.Slides(index).Delete()

        # Move pointer to the last slide
        self._curr_slide += count

    def _add_appendix(self):
        """Adds appendix slides according to the power net list"""
        # Move to first slide of appendix
        start = self._curr_slide + 1
        p_nets = self.power_nets

        self.pptx.Slides(start).Copy()
        sleep(.25)

        # start from 1 to account for init template slide
        for i in range(1, len(p_nets) - 1):
            index = start + i
            self.pptx.Slides.Paste(index)
        
        # Move pointer at start of section to end
        for j in range(0, len(p_nets) - 1):
            shapes = self.pptx.Slides(start + j).Shapes
            for s in shapes:
                if s.HasTextFrame == MSOTRUE:
                    text = s.TextFrame.TextRange.Text[:]
                    if text.startswith("Appendix"):
                        new = text.replace("<i>", str(j + 1))
                        new.replace("<POWER_NET[i]>", p_nets[j][0])
                        s.TextFrame.TextRange.Text = new
                elif s.HasTable == MSOTRUE:
                    text_range = s.Table.Cell(2,2).Shape.TextFrame.TextRange
                    text = text_range.Text[:]
                    new = text.replace("<POWER_NET[i]>", p_nets[j][0])
                    text_range.Text = new

        self.pptx.Slides(start).Delete()
    
    def _build_slides(self, conf_tools):
        self._get_power_nets(conf_tools)
        self._fill_analysis_table()
        self._make_reson_analysis()
        self._add_appendix()

    def build_pptx(self, conf_tools):
        self._make_cover(conf_tools)
        self._copy_slides(conf_tools)
        self._build_slides(conf_tools)
        self._save_report()