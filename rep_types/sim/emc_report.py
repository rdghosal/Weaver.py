from ..simreport import SimulationReport

SIM_TARGETS = 6


class EMCReport(SimulationReport):
    """
    Class for PCB EMC report
    """
    def __init__(self, template, proj_num):
        super().__init__(template, proj_num)
        self.__power_nets = []
        self.__curr_slide = 1 # To store state of slide making

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
    
    def _get_power_nets(self, toc, conf_tools):
        """Reads table of contents (TOC) for pages that need making"""
        tar_pages = toc["sim_targets"]
        # Get page(s) and copy them into report

        count = 0
        for p in tar_pages:
            curr_slide = SIM_TARGETS + count
            conf_tools.pptx.Slides(p).Copy()
            self.pptx.Slides.Paste(curr_slide)
            slide_title = self.pptx.Slides(curr_slide).Shapes(TITLE_NAME).TextFrame.TextRange.Text[:]
            slide_title = slide_title.replace("3.1", "")
            self.pptx.Slides(curr_slide).Shapes(TITLE_NAME).TextFrame.TextRange.Text = slide_title[:]

        sim_tar_table = self._get_table(self.pptx.Slides(SIM_TARGETS).Shapes)
        row = 2 # Initial row for scan
        while True:
            try:
                net_col = 1 # For net names
                resonance_col = 4 # For y/n power resonance analysis
                has_resonance = False 
                net = sim_tar_table.Cell(row, net_col).Shape.TextFrame.TextRange.Text[:]
                if sim_tar_table.Cell(row, resonance_col).Shape.TextFrame.TextRange.Text[:] == u"〇":
                    has_resonance = True
                self.__power_nets.append((net, has_resonance))
                row += 1 # Move down table
            except IndexError:
                # Have reached end of table, so break while loop
                break

        self.__curr_slide += count # Move slide pointer
        return self.power_nets


    def _fill_analysis_table(self):
        """Populates resonance analysis table with power net names"""
        self.__curr_slide += 2 # Move past divider (assumes execution after get_power_nets)
        index = self.__curr_slide        

        # Grab table from slide
        slide = self.pptx.Slide(index)
        table = self._get_table(slide.Shapes)

        row = 2 # init row
        count = 0
        item_num = 1 
        while True:
            try:
                if count % 2 != 0:
                    table.Cell(row, 1).Shape.TextFrame.TextRange.Text = str(item_num)
                    item_num += 1
                new = table.Cell(row, 2).Shape.TextFrame.TextRange.Text.replace("<POWER_NET[i]>", self.power_nets[item_num - 1])
                table.Cell(row, 2).Shape.TextFrame.TextRange.Text = new
                count += 1
            except IndexError:
                break

    def _make_reson_analysis(self):
        """Copy template for resonance analysis and fill in table and title"""
        self.__curr_slide += 3 # move to next (needs better error-proofing)
        index = self.__curr_slide
        shapes = self.pptx.Slide(index).Shapes

        # Exclude init template slide and calc number of times to copy
        num_nets = len(self.power_nets) - 1
        count = 0
        self.pptx.Slides(index).Copy
        # Use filter to only get those nets that need resonance analysis
        p_nets = self.power_nets
        while count < num_nets:
            self.pptx.Slides.Paste(index + count) # Place right after current
            shapes = self.pptx.Slides.Shapes
            for s in shapes:
                if s.HasTextFrame == MSOTRUE:
                    text = s.TextFrame.TextRange.Text[:].lower()
                    if text.startswith("target"):
                        # TODO: use boolean to make sure only needs needing resonance analysis are used
                        new = text.replace("<V[i]>", p_nets[count][1][:] + "V")
                        new = new.replace("<POWER_NET[i]", p_nets[count][0][:])
                        s.TextFrame.TextRange.Text = new
                elif s.HasTable == MSOTRUE:
                    s.Table.Cells(2, 1).Shape.TextFrame.TextRange.Text = p_nets[count]

            # Move to next power net        
            count += 1



            count += 1

        # Move pointer to the next slide 
        self.__curr_slide += count + 1 







        