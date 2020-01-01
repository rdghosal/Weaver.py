import re
from time import sleep
from .. import SimulationReport
from util import TOC, EXEC_SUMM, MSOTRUE, com_error


class SIReport(SimulationReport):
    """
    Class for PCB signal integrity report
    """
    def __init__(self, template, interface, proj_num):
        super().__init__(template, proj_num)
        self.__interface = interface
    
    def __str__(self):
        return f"{self.report_type} Report for {self.interface.name}"

    @property
    def title(self):
        return f"{self.proj_num}\nVerification of Signal Integrity\n{self.interface.name} [Ver.1.0]"

    @property
    def report_type(self):
        return "SI"

    @property
    def interface(self):
        return self.__interface
    
    def _fill_toc(self):
        """Fills in Table of Contents"""
        toc_slide = self.pptx.Slides(TOC)
        toc_table = self._get_table(toc_slide.Shapes)
        # Replace placeholders with Interface.name
        row = 2
        while True:
            try:
                # Grab curr cell contents
                curr_text = toc_table.Cell(row, 1).Shape.TextFrame.TextRange.Text[:]                
                # Reached end of table
                if curr_text == "":
                    break
                # Replace found placeholder
                if curr_text.find("<INTERFACE>") > -1:
                    toc_table.Cell(row, 1).Shape.TextFrame.TextRange.Text = \
                        curr_text.replace("<INTERFACE>", self.interface.name)
                row +=1
            except com_error:
                break

    def _fill_exec_summ(self):
        """Replaces some placeholders in exec summary"""
        exec_summ_slide = self.pptx.Slides(EXEC_SUMM)
        for shape in exec_summ_slide.Shapes:
            if shape.HasTextFrame == MSOTRUE:
                placeholder = "<INTERFACE>"
                curr_text = shape.TextFrame.TextRange.Text[:]
                if curr_text.find(placeholder) > -1:
                    shape.TextFrame.TextRange.Text = curr_text.replace(placeholder, self.interface.name)
    
    def _copy_slides(self, conf_tools):
        """Copies target slides into new report"""
        toc = conf_tools.get_toc()

        # Copy/Paste eye mask slides
        page_ranges = [ toc["eye_mask_judgement"], toc["topology"] ]
        self._curr_slide = 5 # Pasting should start after Methodology slide

        # Copies all eye mask slides and needs author to delete those unneeded
        for pages in page_ranges:
            for i in range(pages[0], pages[1] + 1):
                conf_tools.pptx.Slides(i).Copy()
                sleep(.25)
                self.pptx.Slides.Paste(self._curr_slide)
                self._curr_slide += 1
    
    def _fill_divider(self):
        divider = self.pptx.Slides(self._curr_slide)
        for shape in divider.Shapes:
            if shape.HasTextFrame == MSOTRUE:
                placeholder = "<INTERFACE>"
                curr_text = shape.TextFrame.TextRange.Text[:]
                if curr_text.find(placeholder):
                    shape.TextFrame.TextRange.Text = curr_text.replace(placeholder, self.interface.name)

    def _fill_results_table(self):
        """Fills Results table with signal info"""
        results_table_slide = self.pptx.Slides(self._curr_slide)
        results_table = self._get_table(results_table_slide.Shapes)
        
        row = 5
        is_staggered = True # To mark whether on a staggered row (rows within row)

        # Add extra rows if necessary
        diff = len(self.interface.signals) - 4
        if diff > 0:
            for i in range(diff):
                results_table.Rows.Add()

        # Iterate over signals and fill in table
        for i, signal in enumerate(self.interface.signals):
            # Additional rows are not staggered format
            if i > 4:
                is_staggered = False
            # Text for cell in each column
            text = {
                1: signal.name,
                2: "\n".join(signal.frequency),
                3: "\n".join([ signal.driver.ibis_model, signal.driver.buffer_model ]),
                4: "\n".join([ signal.receiver.ibis_model, signal.receiver.buffer_model ]),
                5: [ signal.pvt[0], signal.pvt[1] ]
            }

            for col in range(1, 6):
                if col == 5:
                    tar_text = " ".join(text[col]) if not is_staggered else text[col][0]
                    results_table.Cell(row, col).Shape.TextFrame.TextRange.Text = tar_text
                    if is_staggered:
                        row += 1
                        results_table.Cell(row, col).Shape.TextFrame.TextRange.Text = text[col][1]
                else:
                    results_table.Cell(row, col).Shape.TextFrame.TextRange.Text = text[col]

    def _replace_placeholder(self, curr_text, signal_count):
        """Searches text for a potential placeholder and returns new text"""
        placeholders = {
            "<INTERFACE>": self.interface.name,
            "<SIGNAL>": self.interface.signals[signal_count].name,
            "<FREQ>": " ".join(self.interface.signals[signal_count].frequency),
            "<DRIVER_IBS>": self.interface.signals[signal_count].driver.ibis_model,
            "<DRIVER_MODEL>": self.interface.signals[signal_count].driver.buffer_model,
            "<RECEIVER_IBS>": self.interface.signals[signal_count].receiver.ibis_model,
            "<RECEIVER_MODEL>": self.interface.signals[signal_count].receiver.buffer_model
        }

        match = re.search(r".*(<\w+>).*", curr_text)
        if match: 
            match = match.group(1) # Get capture group
            if match in placeholders.keys():
                return curr_text.replace(match, placeholders[match])
            else:
                return ""

    def _build_slides(self):
        self.pptx.Slides(self._curr_slide).Copy() # Copy template slide
        sleep(.25)
        diff = len(self.interface.signals) - 1 # Accounts for 1 template slide
        slide_ptr = self._curr_slide # Memoize first slide index
        signal_count = 0
        if diff > 0:
            for _ in range(diff):
                self._curr_slide += 1
                self.pptx.Slides.Paste(self._curr_slide)

        while slide_ptr <= self._curr_slide:
            for shape in self.pptx.Slides(slide_ptr).Shapes:
                if shape.HasTextFrame == MSOTRUE:
                    curr_text = shape.TextFrame.TextRange.Text[:]
                    if curr_text:
                        new_text = self._replace_placeholder(curr_text, signal_count)
                        if new_text: shape.TextFrame.TextRange.Text = new_text 
                elif shape.HasTable == MSOTRUE:
                    # Check first header cell
                    if shape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text == "Item":
                        tar_cells = [(2,2), (4,1), (4,2), (4,3), (5,1), (5,2), (5,3)]
                        for cell in tar_cells:
                            curr_text = shape.Table.Cell(cell[0], cell[1]).Shape.TextFrame.TextRange.Text[:]
                            shape.Table.Cell(cell[0], cell[1]).Shape.TextFrame.TextRange.Text = self._replace_placeholder(curr_text, signal_count)

            slide_ptr += 1
            signal_count += 1
        
    def build_pptx(self, conf_tools):
        # Name composed of more than one word
        # and is ignored for being a special case report
        if self.interface.name.find(" ") > -1:
            print("Ignoring build for interface") 
            print(f"{self.interface.name} in {self.proj_num}")
            return

        self._make_cover(conf_tools)
        self._fill_toc()
        self._fill_exec_summ()
        self._copy_slides(conf_tools)
        self._fill_divider()
        self._curr_slide += 1 # Move to Results table
        self._fill_results_table()
        self._curr_slide += 1
        self._build_slides()
        self._save_report()
        