"""
LabelSpec: the in-memory representation of a preset, as loaded from a preset
.json file (see preset_store.py) and consumed by main.py/label_format.py to
generate a label document.

Preset .json files may carry extra bookkeeping keys that aren't part of the
generation logic (e.g. "name", "preset_id", "saved_headers") - those are
accepted and ignored rather than raising an error, so new metadata fields can
be added to the preset format without breaking old code that loads them.
"""


class LabelSpec:
    def __init__(
        self,
        *,
        presettype=None,
        copiesperlabel=None,
        multi_copiesperlabel=None,
        textboxformatinput=None,
        labeltemplate=None,
        fontname=None,
        fontsize=None,
        text_alignment="Center",
        outputfilenameprefix=None,
        output_add_date=None,
        outputformat=None,
        truncation_indices=None,
        text_box_input=None,
        identical_or_incremental=None,
        color_theme=None,
        ui_layout=None,
        partialsheet=None,
        row_start=None,
        row_end=None,
        col_start=None,
        col_end=None,
        pages_of_labels=1,
        date_format=None,
        sample_filename=None,
        remove_duplicates=None,
        **_ignored_extra_fields,
    ):
        self.presettype = presettype  # "Text" or "File"
        self.copiesperlabel = copiesperlabel
        self.multi_copiesperlabel = multi_copiesperlabel
        self.textboxformatinput = textboxformatinput
        self.labeltemplate = labeltemplate  # key into label_templates.label_templates
        self.fontname = fontname
        self.fontsize = fontsize
        self.alignment = text_alignment  # "Left" / "Center" / "Right"
        self.outputfilenameprefix = outputfilenameprefix
        self.output_add_date = output_add_date
        self.outputformat = outputformat
        self.truncation_indices = truncation_indices
        self.text_box_input = text_box_input
        self.identical_or_incremental = identical_or_incremental  # "Identical" or "Incremental", for Text presets
        self.color_theme = color_theme
        self.ui_layout = ui_layout if ui_layout is not None else {}
        self.partialsheet = partialsheet
        self.row_start = row_start
        self.row_end = row_end
        self.col_start = col_start
        self.col_end = col_end
        self.pages_of_labels = pages_of_labels
        self.date_format = date_format
        self.sample_filename = sample_filename
        self.remove_duplicates = remove_duplicates
