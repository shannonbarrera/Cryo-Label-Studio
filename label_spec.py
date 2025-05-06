
class LabelSpec:
    def __init__(self, **kwargs):
        self.inputtype = kwargs.get("inputtype")
        self.copiesperlabel = kwargs.get("copiesperlabel")
        self.tablecoords = kwargs.get("tablecoords")
        self.textboxformatinput = kwargs.get("textboxformatinput")
        self.labeltemplatepath = kwargs.get("labeltemplatepath")
        self.labelsheetlayouttype = kwargs.get("labelsheetlayouttype")
        self.fontname = kwargs.get("fontname")
        self.fontsize = kwargs.get("fontsize")
        self.outputfilenameprefix = kwargs.get("outputfilenameprefix")
        self.outputformat = kwargs.get("outputformat")
        self.output_file_path = kwargs.get("output_file_path")
        self.input_file_path = kwargs.get("input_file_path")
        self.truncation_indices = kwargs.get("truncation_indices")
        self.text_box_input = kwargs.get("text_box_input")
        self.identical_or_incremental = kwargs.get("identical_or_incremental")
