import sys

from label_templates import label_templates

from data_extract import (
    get_data_list_csv,
    get_data_list_xlsx
)

from data_process import (
    truncate_data
)

from file_io import (
    get_template,
    get_file_path,
    save_file
)

from label_format import (
    generate_blank_docx_template,
    get_max_labels_per_page,
    format_labels_single,
    format_labels_multi,
    format_labels_identical,
    format_labels_incremental,
    combine_docs,
    get_layout_from_spec
)

from label_spec import LabelSpec


def main(spec: LabelSpec, input_file_path=None, output_file_path=None, text_box_input=None):
    labeltemplate = generate_blank_docx_template(spec.labeltemplate)
    row_indices, column_indices = get_layout_from_spec(spec)
    
    multi_pages = False
    final_doc = None

    if spec.presettype == "File":
        # Load data from file based on extension
        if input_file_path.lower().endswith(".csv"):
            data_list = get_data_list_csv(input_file_path, spec.textboxformatinput)
        elif input_file_path.lower().endswith((".xls", ".xlsx")):
            data_list = get_data_list_xlsx(input_file_path, spec.textboxformatinput)
        else:
            raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")

        # Optional truncation
        if spec.truncation_indices:
            data_list = truncate_data(data_list, spec.truncation_indices)
        
        format_function = format_labels_multi if spec.copiesperlabel > 1 else format_labels_single
        max_labels_per_page = get_max_labels_per_page(spec, labeltemplate, spec.copiesperlabel)
        pages = [
            data_list[i : i + max_labels_per_page]
            for i in range(0, len(data_list), max_labels_per_page)
        ]
        print("60")
        final_doc = format_function(pages[0], labeltemplate, row_indices, column_indices, spec.copiesperlabel, spec.textboxformatinput, spec.fontname, spec.fontsize)
        print("62")
        for page in pages[1:]:
            next_doc = format_function(page, labeltemplate, row_indices, column_indices, spec.copiesperlabel, spec.textboxformatinput, spec.fontname, spec.fontsize)
            final_doc = combine_docs(final_doc, next_doc)

        print("67")
    elif spec.presettype == "Text":
        logic = spec.identical_or_incremental
        print("logic")
        if logic == "identical":
            final_doc = format_labels_identical(text_box_input, labeltemplate, row_indices, column_indices, spec.fontname, spec.fontsize)
        elif logic == "incremental":
            final_doc = format_labels_incremental(text_box_input, labeltemplate, row_indices, column_indices, spec.fontname, spec.fontsize)
        elif logic == "serial":
            try:
                count = int(spec.labels_perserial)
            except (TypeError, ValueError):
                count = 1
            num_serials = 100  # or make this configurable
            labels = []
            for i in range(1, num_serials + 1):
                serial = f"{i:03}"
                for _ in range(count):
                    label = spec.textboxformatinput.replace("{serial}", serial)
                    labels.append(label)
            final_doc = format_labels_identical("\\n".join(labels), labeltemplate, row_indices, column_indices, spec.fontname, spec.fontsize)

    else:
        raise ValueError("Invalid presettype: must be 'Text' or 'File'")

    # Determine save location
    save_path = get_file_path(
        output_file_path or spec.output_file_path,
        spec.outputfilenameprefix,
        spec.outputformat
    )
    print(save_path)
    save_file(save_path, final_doc)

    
if __name__ == "__main__":
    """
    Entry point of the script when run from the command line. Ensures that the user provides two arguments:
    the excel sheet location and the quantity of labels.
    """
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print("Please provide 1 argument")