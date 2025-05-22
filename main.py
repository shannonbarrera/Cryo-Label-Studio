import sys
import re
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
    get_row_and_column_indices,
    get_first_page_row_indices,
    get_first_page_col_indices,
    get_max_labels_per_page,
    get_max_labels_first_page,
    paginate_labels,
    format_labels_firstpage_fromfile,
    format_labels_single,
    format_labels_multi,
    format_labels_identical,
    format_labels_incremental,
    combine_docs,
)

from label_spec import LabelSpec
from docx import Document

def main(spec: LabelSpec, input_file_path=None, output_file_path=None, text_box_input=None):
    template_meta = label_templates[spec.labeltemplate]
    templatepath = template_meta["template_path"]
    labeltemplateexample = Document(templatepath)
    table_format = template_meta["table_format"]
    start_row = getattr(spec, "row_start", 1)
    end_row = getattr(spec, "row_end", template_meta.get("labels_down", 99))
    start_col = getattr(spec, "col_start", 1)
    end_col = getattr(spec, "col_end", template_meta.get("labels_across", 99))


    row_indices, column_indices = get_row_and_column_indices(templatepath, table_format)

    first_page_row_indices = get_first_page_row_indices(start_row, end_row, row_indices)
    print(first_page_row_indices)
    first_page_first_row_col_indices, first_page_last_row_col_indices = get_first_page_col_indices(start_col, end_col, start_row, end_row, column_indices)

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
        max_labels_per_page = get_max_labels_per_page(spec, templatepath, table_format)

        first_page_max_labels = get_max_labels_first_page(first_page_row_indices, column_indices, first_page_first_row_col_indices, first_page_last_row_col_indices)
        first_page, otherpages = paginate_labels(first_page_max_labels, max_labels_per_page, data_list, spec.copiesperlabel)
        pages = [first_page]

        pages = pages + otherpages

        final_doc = format_labels_firstpage_fromfile(pages[0], templatepath, first_page_row_indices, column_indices, first_page_first_row_col_indices, first_page_last_row_col_indices, spec)

        for page in pages[1:]:
            next_doc = format_labels_single(page, templatepath, row_indices, column_indices, spec)
            final_doc = combine_docs(final_doc, next_doc)

        save_file(output_file_path, final_doc)
        return
    
    elif spec.presettype == "Text":
        logic = spec.identical_or_incremental
        print(f"Logic: {logic}")

        if logic == "Identical":
            data_list = get_data_list_fromtext(text_box_input, logic)
            print(data_list)
            final_doc = format_labels_identical(
                text_box_input,
                templatepath,
                row_indices,
                column_indices,
                spec.fontname,
                spec.fontsize
            )

        elif logic == "Incremental":

            match = re.match(r"([A-Za-z0-9\-_]*?)(\d+)$", text_box_input)
            if not match:
                messagebox.showerror("Error", "Starting serial must end in a number (e.g., AB-001)")
                return

            prefix, start_num = match.groups()
            num_digits = len(start_num)
            start = int(start_num)

            try:
                count = int(spec.copiesperlabel)
            except (TypeError, ValueError):
                count = 1
            print(count)
            first_page_max_labels = get_max_labels_first_page(
                first_page_row_indices,
                column_indices,
                first_page_first_row_col_indices,
                first_page_last_row_col_indices
            )

            num_serials = first_page_max_labels // count
            labels = []

            for i in range(start, start + num_serials):
                serial_num = f"{i:0{num_digits}d}"
                serial = f"{prefix}{serial_num}"
                for _ in range(count):
                    label = spec.textboxformatinput.replace("{LABEL_TEXT}", serial)
                    labels.append([label])
            print(labels)


            final_doc = format_labels_firstpage_fromfile(labels, templatepath, first_page_row_indices, column_indices, first_page_first_row_col_indices, first_page_last_row_col_indices, spec)

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