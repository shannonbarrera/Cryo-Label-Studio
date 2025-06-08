"""
Main module for label generation using CryoPop Label Studio presets.

This script processes label specifications (via LabelSpec) and generates formatted Word documents
based on either text-based or file-based input. It handles template selection, layout calculation,
pagination, serial number generation, and document formatting.

Supported input types:
- File presets (.csv or .xlsx input)
- Text presets with either identical or incremental logic

The resulting document is saved to the specified output path.
"""

import sys
import re
from label_templates import label_templates
from data_extract import get_data_list_csv, get_data_list_xlsx
from file_io import get_file_path, save_file, get_template
from label_format import (
    get_row_and_column_indices,
    get_first_page_row_indices,
    get_first_page_col_indices,
    get_max_labels_per_page,
    get_max_labels_first_page,
    paginate_labels,
    format_labels_page,
    combine_docs,
)
from label_spec import LabelSpec
from docx import Document


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
    templatepath = template_meta["template_path"]
    labeltemplateexample = Document(templatepath)
    table_format = template_meta["table_format"]
    start_row = getattr(spec, "row_start", 1)
    end_row = getattr(spec, "row_end", template_meta.get("labels_down", 99))
    start_col = getattr(spec, "col_start", 1)
    end_col = getattr(spec, "col_end", template_meta.get("labels_across", 99))

    row_indices, column_indices = get_row_and_column_indices(templatepath, table_format)

    if spec.partialsheet == True:
        first_page_row_indices = get_first_page_row_indices(
            start_row, end_row, row_indices
        )

        first_page_first_row_col_indices, first_page_last_row_col_indices = (
            get_first_page_col_indices(
                start_col, end_col, start_row, end_row, column_indices
            )
        )

    else:
        first_page_row_indices = row_indices
        first_page_first_row_col_indices = column_indices
        first_page_last_row_col_indices = column_indices

    multi_pages = False
    final_doc = None



    if spec.presettype == "File":
        # Load data from file based on extension
        if input_file_path.lower().endswith(".csv"):
            data_list = get_data_list_csv(input_file_path, spec.textboxformatinput, spec.date_format)
        elif input_file_path.lower().endswith((".xls", ".xlsx")):
            data_list = get_data_list_xlsx(input_file_path, spec.textboxformatinput, spec.date_format)
        else:
            raise ValueError(
                "Unsupported file type. Please upload a .csv or .xlsx file."
            )

        max_labels_per_page = get_max_labels_per_page(spec, templatepath, table_format)

        first_page_max_labels = get_max_labels_first_page(
            first_page_row_indices,
            column_indices,
            first_page_first_row_col_indices,
            first_page_last_row_col_indices,
        )
        first_page, otherpages = paginate_labels(
            first_page_max_labels, max_labels_per_page, data_list, spec.copiesperlabel
        )
        pages = [first_page]

        pages = pages + otherpages

        for i, page in enumerate(pages):
            is_last = (i == len(pages) - 1)
            formatted_page = format_labels_page(
                page,
                templatepath,
                first_page_row_indices if i == 0 else row_indices,
                column_indices,
                first_page_first_row_col_indices if i == 0 else column_indices,
                first_page_last_row_col_indices if i == 0 else column_indices,
                spec,
                is_last_page=is_last  # pass this flag!
            )

            if final_doc is None:
                final_doc = formatted_page
            else:
                final_doc = combine_docs(final_doc, formatted_page)

                save_file(output_file_path, final_doc)
                return

    elif spec.presettype == "Text":
        logic = spec.identical_or_incremental

        if logic == "Identical":
            labeltext = text_box_input

            try:
                count = int(spec.copiesperlabel)
            except (TypeError, ValueError):
                # Fill the page if copiesperlabel is blank or invalid
                count = get_max_labels_first_page(
                    first_page_row_indices,
                    column_indices,
                    first_page_first_row_col_indices,
                    first_page_last_row_col_indices
                )


            data_list = [labeltext] * count

            first_page_max_labels = get_max_labels_first_page(
                first_page_row_indices,
                column_indices,
                first_page_first_row_col_indices,
                first_page_last_row_col_indices,
            )
            max_labels_per_page = get_max_labels_per_page(spec, templatepath, table_format)
            
            pages = [*paginate_labels(first_page_max_labels, max_labels_per_page, data_list, 1)]

            final_doc = format_labels_page(
                pages[0],
                templatepath,
                first_page_row_indices,
                column_indices,
                first_page_first_row_col_indices,
                first_page_last_row_col_indices,
                spec,
            )

            for page in pages[1:]:
                if len(page) > 0:
                    next_doc = format_labels_page(
                        page,
                        templatepath,
                        row_indices,
                        column_indices,
                        column_indices,
                        column_indices,
                        spec,
                    )
                    final_doc = combine_docs(final_doc, next_doc)


        elif logic == "Incremental":
            num_pages = spec.pages_of_labels
            match = re.match(r"([A-Za-z0-9\-_]*?)(\d+)$", text_box_input)
            if not match:
                messagebox.showerror(
                    "Error", "Starting serial must end in a number (e.g., AB-001)"
                )
                return

            prefix, start_num = match.groups()
            num_digits = len(start_num)
            start = int(start_num)

            try:
                count = int(spec.copiesperlabel)
            except (TypeError, ValueError):
                count = 1

            first_page_max_labels = get_max_labels_first_page(
                first_page_row_indices,
                column_indices,
                first_page_first_row_col_indices,
                first_page_last_row_col_indices,
            )

            max_labels_per_page = get_max_labels_per_page(
                spec, templatepath, table_format
            )
            labelcount_additional_pages = max_labels_per_page * (num_pages - 1)

            num_serials = (first_page_max_labels + labelcount_additional_pages) // count
            data_list = []

            for i in range(start, start + num_serials):
                serial_num = f"{i:0{num_digits}d}"
                serial = f"{prefix}{serial_num}"
                for _ in range(count):
                    label = serial
                    data_list.append([label])

            firstpage, otherpages = paginate_labels(
                first_page_max_labels, max_labels_per_page, data_list, 1
            )

            pages = [firstpage]

            pages = pages + otherpages

            for i, page in enumerate(pages):
                is_last = (i == len(pages) - 1)
                formatted_page = format_labels_page(
                    page,
                    templatepath,
                    first_page_row_indices if i == 0 else row_indices,
                    column_indices,
                    first_page_first_row_col_indices if i == 0 else column_indices,
                    first_page_last_row_col_indices if i == 0 else column_indices,
                    spec,
                    is_last_page=is_last  # pass this flag!
                )

                if final_doc is None:
                    final_doc = formatted_page
                else:
                    final_doc = combine_docs(final_doc, formatted_page)


    else:
        raise ValueError("Invalid presettype: must be 'Text' or 'File'")

    save_file(output_file_path, final_doc)


if __name__ == "__main__":
    """
    Entry point of the script when run from the command line. Ensures that the user provides two arguments:
    the excel sheet location and the quantity of labels.
    """
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print("Please provide 1 argument")
