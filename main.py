import sys

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
    get_max_labels_per_page,
    format_labels_single,
    format_labels_multi,
    format_labels_identical,
    format_labels_incremental,
    combine_docs
)



from label_spec import LabelSpec

def main(spec: LabelSpec):
    """
    Generates formatted labels from a CSV or Excel file and saves them to a Word document.

    This function handles the entire label creation pipeline:
    - Loads a label template.
    - Extracts and optionally truncates data from the input file.
    - Formats the data into label layouts according to a specified template and layout type.
    - Applies formatting (font, size, etc.) based on user input.
    - Handles multi-page documents if the number of labels exceeds one page.
    - Saves the final document to the specified output path.

    Args:
        kwargs (dict): A dictionary containing the following keys:
            - inputtype (str): The type of input file ('CSV' or 'XLSX').
            - copiesperlabel (int): Number of identical labels to generate for each data row.
            - tablecoords (tuple): Starting row/column coordinates for the data table in the Excel file.
            - textboxformatinput (list): List of format instructions for each text box.
            - labeltemplatepath (str): Path to the label template (.docx file).
            - labelsheetlayouttype (str): Label sheet layout type (e.g. 'single', 'multi').
            - fontname (str): Font name to use in label text.
            - fontsize (int): Font size for label text.
            - outputfilenameprefix (str): Prefix to apply to the output filename.
            - outputformat (str): File format for the output document (e.g., 'docx').
            - output_file_path (str): Path to save the output file.
            - input_file_path (str): Path to the input CSV or Excel file.
            - truncation_indices (list): Indices for truncating data fields before formatting.
            - text_box_input (str): The user-entered label text.

    Returns:
        None
    """

    inputtype = spec.inputtype
    copiesperlabel = spec.copiesperlabel
    tablecoords = spec.tablecoords
    textboxformatinput = spec.textboxformatinput
    labeltemplatepath = spec.labeltemplatepath
    labelsheetlayouttype = spec.labelsheetlayouttype
    fontname = spec.fontname
    fontsize = spec.fontsize
    outputfilenameprefix = spec.outputfilenameprefix
    outputformat = spec.outputformat
    output_file_path = spec.output_file_path
    input_file_path = spec.input_file_path
    truncation_indices = spec.truncation_indices
    text_box_input = spec.text_box_input
    identical_or_incremental = spec.identical_or_incremental

    ''' Need some function which normalizes the textboxformatinput in reference to the table.  When it 
    comes in it's in reference to the entire excel sheet, which the table doesn't necessarily start at 0,0
    '''
    
    # Get the label template file which is stored as a data file
    labeltemplate = get_template(labeltemplatepath)

    # Get the row and column indices and the number of label cells per page
    # These next two functions need to be updated to handle all labelsheetlayouttypes
    row_indices, column_indices = get_row_and_column_indices(labeltemplate, labelsheetlayouttype)
    max_labels_per_page = get_max_labels_per_page(labeltemplate, labelsheetlayouttype, copiesperlabel)

    # Get the data_list from the excel or csv file if inputtype is one of those
    
    if inputtype == "CSV":
        data_list = get_data_list_csv(input_file_path, tablecoords, textboxformatinput)
        multi_pages = True
    if inputtype == "XLSX":
        data_list = get_data_list_xlsx(input_file_path, tablecoords, textboxformatinput)
        multi_pages = True
    if truncation_indices:
        data_list = truncate_data(data_list, truncation_indices)

    # Get the data_list if inputtype is "textbox"
    if inputtype == "textbox":
        multi_pages = False
        if identical_or_incremental == "identical":
            final_doc = format_labels_identical(text_box_input, labeltemplate, row_indices, column_indices, fontname, fontsize)
        if identical_or_incremental == "incremental":
            final_doc = format_labels_incremental(text_box_input, labeltemplate, row_indices, column_indices, fontname, fontsize)
    
    # Split data into chunks of max_labels_per_page
    if multi_pages == True:
        if copiesperlabel == 1:
            format_function = format_labels_single
        else:
            format_function = format_labels_multi
        pages = [
            data_list[i : i + max_labels_per_page]
            for i in range(0, len(data_list), max_labels_per_page)
        ]
        # Process first page
        final_doc = format_function(pages[0], labeltemplate, row_indices, column_indices, copiesperlabel, textboxformatinput, fontname, fontsize,)

        # Process remaining pages and append to final_doc
        for page in pages[1:]:
            next_page_doc = format_function(page, labeltemplate, row_indices, column_indices, copiesperlabel, textboxformatinput, fontname, fontsize)
            final_doc = combine_docs(final_doc, next_page_doc)

    

    outputfilelocation = get_file_path(output_file_path, outputfilenameprefix, outputformat)
    print(outputfilelocation)
    save_file(outputfilelocation, final_doc)
    return

    
if __name__ == "__main__":
    """
    Entry point of the script when run from the command line. Ensures that the user provides two arguments:
    the excel sheet location and the quantity of labels.
    """
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print("Please provide 1 argument")