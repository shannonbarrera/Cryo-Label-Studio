
class LabelSpec:
    def __init__(self, **kwargs):
        self.presettype = kwargs.get("presettype")
        self.inputtype = kwargs.get("inputtype")
        self.copiesperlabel = kwargs.get("copiesperlabel")
        self.textboxformatinput = kwargs.get("textboxformatinput")
        self.labeltemplatepath = kwargs.get("labeltemplatepath")
        self.fontname = kwargs.get("fontname")
        self.fontsize = kwargs.get("fontsize")
        self.outputfilenameprefix = kwargs.get("outputfilenameprefix")
        self.outputformat = kwargs.get("outputformat")
        self.truncation_indices = kwargs.get("truncation_indices")
        self.text_box_input = kwargs.get("text_box_input")
        self.identical_or_incremental = kwargs.get("identical_or_incremental")
        self.labels_perserial = kwargs.get("labels_perserial")

