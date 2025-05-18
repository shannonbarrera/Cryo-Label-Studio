
class LabelSpec:
    def __init__(self, **kwargs):
        self.presettype = kwargs.get("presettype")
        self.copiesperlabel = kwargs.get("copiesperlabel")
        self.textboxformatinput = kwargs.get("textboxformatinput")
        self.labeltemplate = kwargs.get("labeltemplate")
        self.fontname = kwargs.get("fontname")
        self.fontsize = kwargs.get("fontsize")
        self.outputfilenameprefix = kwargs.get("outputfilenameprefix")
        self.output_add_date = kwargs.get("output_add_date")
        self.outputformat = kwargs.get("outputformat")
        self.truncation_indices = kwargs.get("truncation_indices")
        self.text_box_input = kwargs.get("text_box_input")
        self.identical_or_incremental = kwargs.get("identical_or_incremental")
        self.labels_perserial = kwargs.get("labels_perserial")
        self.color_theme = kwargs.get("color_theme")
        self.ui_layout = kwargs.get("ui_layout", {})
        self.partialsheet = kwargs.get("partialsheet")

