"""
Main module for label generation using Cryo Label Studio presets.

This script processes label specifications (via LabelSpec) and generates formatted Word documents
based on either text-based or file-based input. It handles template selection, layout calculation,
pagination, serial number generation, and document formatting.

Supported input types:
- File presets (.csv or .xlsx input)
- Text presets with either identical or incremental logic

The resulting document is saved to the specified output path.
"""

import re
from label_templates import label_templates
from data_extract import get_data_list_csv, get_data_list_xlsx, remove_duplicate_labels
from file_io import save_file
from label_format import smart_wrap_label_text
from label_spec import LabelSpec
from data_process import estimate_max_chars
from page_layout import PageLayout


def _load_file_data_list(spec, input_file_path):
    if input_file_path.lower().endswith(".csv"):
        data_list = get_data_list_csv(input_file_path, spec.textboxformatinput, spec.date_format)
    elif input_file_path.lower().endswith((".xls", ".xlsx")):
        data_list = get_data_list_xlsx(input_file_path, spec.textboxformatinput, spec.date_format)
    else:
        raise ValueError(
            "Unsupported file type. Please upload a .csv or .xlsx file."
        )

    if spec.remove_duplicates == True:
        data_list = remove_duplicate_labels(data_list)

    return data_list


def _build_incremental_data_list(spec, text_box_input, layout, template_meta):
    match = re.match(r"([A-Za-z0-9\-_]*?)(\d+)$", text_box_input)
    if not match:
        raise ValueError(
            f"Could not parse a serial number from '{text_box_input}'. "
            "Expected an optional prefix followed by digits (e.g. 'ab-0001')."
        )

    prefix, start_num = match.groups()
    num_digits = len(start_num)
    start = int(start_num)

    try:
        count = int(spec.copiesperlabel)
    except (TypeError, ValueError):
        count = 1

    num_pages = spec.pages_of_labels
    labelcount_additional_pages = layout.max_labels_per_page * (num_pages - 1)
    num_serials = (layout.first_page_max_labels + labelcount_additional_pages) // count

    max_chars = estimate_max_chars(template_meta["label_width"], float(spec.fontsize), "Arial")

    data_list = []
    for i in range(start, start + num_serials):
        serial = f"{prefix}{i:0{num_digits}d}"
        label = smart_wrap_label_text(serial, max_chars, prefix)
        for _ in range(count):
            data_list.append([label])

    return data_list


def main(
    spec: LabelSpec, input_file_path=None, output_file_path=None, text_box_input=None
):
    """
    Generates formatted labels based on the provided LabelSpec and input data.

    This function supports both file-based and text-based label generation:
    - File presets extract data from CSV/XLSX and populate labels using a format string.
    - Text presets either repeat a static value ("Identical") or increment serials ("Incremental").

    Args:
        spec (LabelSpec): Preset specification defining layout, format, and behavior.
        input_file_path (str, optional): Path to the CSV or XLSX input file for 'File' presets.
        output_file_path (str, optional): Path to save the generated Word document.
        text_box_input (str, optional): Text or serial prefix for 'Text' presets.

    Raises:
        ValueError: If the input file type is unsupported or the preset type is invalid.
        Exception: For issues during data parsing, formatting, or saving.

    Returns:
        None. The final document is saved to disk.
    """

    template_meta = label_templates[spec.labeltemplate]
    layout = PageLayout(spec, template_meta)

    if spec.presettype == "File":
        data_list = _load_file_data_list(spec, input_file_path)
        copiesperlabel = spec.copiesperlabel

    elif spec.presettype == "Text":
        logic = spec.identical_or_incremental

        if logic == "Identical":
            try:
                count = int(spec.copiesperlabel)
            except (TypeError, ValueError):
                # Fill the page if copiesperlabel is blank or invalid
                count = layout.first_page_max_labels

            data_list = [text_box_input] * count
            copiesperlabel = 1

        elif logic == "Incremental":
            data_list = _build_incremental_data_list(spec, text_box_input, layout, template_meta)
            copiesperlabel = 1

        else:
            raise ValueError(
                "Invalid identical_or_incremental: must be 'Identical' or 'Incremental'"
            )

    else:
        raise ValueError("Invalid presettype: must be 'Text' or 'File'")

    final_doc = layout.build_document(data_list, copiesperlabel, spec)
    save_file(output_file_path, final_doc)


if __name__ == "__main__":
    # main() requires a LabelSpec, not a raw path, so there's no meaningful way
    # to drive it from bare command-line arguments. Run gui.py instead.
    print("Cryo Label Studio is a GUI app. Run 'python gui.py' to launch it.")
