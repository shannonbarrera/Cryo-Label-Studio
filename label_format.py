import sys
import re
import copy
from docx.oxml import OxmlElement
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from datetime import datetime
from label_templates import label_templates
from docx.oxml.ns import qn
from docxcompose.composer import Composer
import math


def get_row_and_column_indices(templatepath, table_format):
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]
    if table_format == "checkerboard":
        row_indices = [i for i in range(len(table.rows)) if i % 2 == 0]
        col_indices = [j for j in range(len(table.columns)) if j % 2 == 0]
    
    if table_format == "LSL stripes":
        row_indices = [i for i in range(len(table.rows))]
        col_indices = [j for j in range(len(table.columns)) if j % 2 == 0]

    return row_indices, col_indices

def get_first_page_row_indices(start_row, end_row, row_indices):
    first_page_row_indices = []
    for i in range(len(row_indices)):
        if i >= start_row-1 and i <= end_row-1:
            first_page_row_indices.append(row_indices[i])
    print(first_page_row_indices)
    return first_page_row_indices

def get_first_page_col_indices(start_col, end_col, col_indices):
    first_page_first_row_col_indices = []
    first_page_last_row_col_indices = []
    if end_col - start_col == 0:
        for i in range(len(col_indices)):
            if i >= start_col-1 and i <= end_col -1:
                first_page_first_row_col_indices.append(col_indices[i])
        return first_page_first_row_col_indices, first_page_last_row_col_indices

    for i in range(len(col_indices)):
        if i >= start_col-1:
            first_page_first_row_col_indices.append(col_indices[i])

    for i in range(len(col_indices)):
        if i <= end_col-1:
            first_page_last_row_col_indices.append(col_indices[i])

    if end_col - start_col == 0:
        first_page_first_row_col_indices = []
        for i in range(len(col_indices)):
            if i >= start_col-1:
                first_page_first_row_col_indices.append(col_indices[i])

    return first_page_first_row_col_indices, first_page_last_row_col_indices

    


def get_max_labels_per_page(spec, templatepath, table_format):
    """
    Calculates how many label entries can fit per page.

    Args:
        labeltemplate (str): Path to label template file.
        labelsheetlayouttype (str): Layout type of the label sheet.
        copiesperlabel (int): Number of repeated labels per entry.

    Returns:
        int: Maximum number of unique label entries per page.
    """

    row_indices, column_indices = get_row_and_column_indices(templatepath, table_format)
    total_cells = len(row_indices) * len(column_indices)
    return total_cells 

def get_max_labels_first_page(first_page_row_indices, column_indices, first_page_first_row_col_indices, first_page_last_row_col_indices):
    labels_first_page_first_last_row = len(first_page_first_row_col_indices) + len(first_page_last_row_col_indices)
    labels_first_page_middle_rows = (len(first_page_row_indices) - 2) * len(column_indices)
    total_cells = labels_first_page_first_last_row + labels_first_page_middle_rows
    return total_cells 

def paginate_labels(first_page_max_labels, max_labels_per_page, data_list, copiesperlabel):
    """
    Distributes labels across multiple pages according to page capacity constraints.
    
    This function handles the pagination of labels, accounting for different capacities
    on the first page versus subsequent pages, and managing the replication of each data item
    according to the specified copies per label.
    
    Parameters:
    -----------
    first_page_max_labels : int
        Maximum number of individual labels that can fit on the first page.
    max_labels_per_page : int
        Maximum number of individual labels that can fit on each subsequent page.
    data_list : list
        List of data items to be formatted as labels.
    copiesperlabel : int
        Number of times each data item should be repeated (copies of the same label).
    
    Returns:
    --------
    tuple
        A tuple containing:
        - firstpage: List of labels for the first page
        - pages: List of lists, where each inner list contains labels for a subsequent page
          (empty if all labels fit on the first page)
    """
    # Calculate total number of individual labels needed
    total_labels = len(data_list) * copiesperlabel

    # Case 1: All labels fit on the first page
    if first_page_max_labels > len(data_list) * copiesperlabel:
        num_pages = 1
        firstpage = []
        pages = []  # Will remain empty since all labels fit on first page
        remainder = 0
        data_count = 0
        
        # Build the first page, repeating each data item according to copiesperlabel
        while len(firstpage) < len(data_list):
            for i in range(copiesperlabel):
                firstpage.append(data_list[data_count])
                remainder = copiesperlabel - (i + 1)
                if remainder == 0:
                    # Move to next data item after adding all copies
                    data_count += 1
  
    # Case 2: Labels span multiple pages
    else:
        # Calculate total number of pages needed
        num_pages = math.ceil((total_labels - first_page_max_labels) / max_labels_per_page) + 1
        firstpage = []

        # Calculate how many data items can fit on first page
        first_page_max_data = first_page_max_labels // copiesperlabel
        
        # Adjust if there's room for part of another data item's copies
        if first_page_max_data * copiesperlabel < first_page_max_labels:
            first_page_max_data += 1

        # Build the first page with copies of each data item
        for item in data_list[:first_page_max_data]:
            for i in range(copiesperlabel):
                firstpage.append(item)
        
        # Handle overflow if we added too many labels to the first page
        if len(firstpage) > first_page_max_labels:
            # Store overflow labels as "hangover" for the next page
            hangover = [firstpage[first_page_max_labels:]]
        else: 
            hangover = []
        
        # Trim first page to maximum capacity
        firstpage = firstpage[:first_page_max_labels]
        
        # Combine hangover with remaining unprocessed data items
        remaining = hangover + data_list[first_page_max_data:]
        pages = []
        remainingindex = 0
        page = []
        
        # Process subsequent pages
        for i in range(1, num_pages):
            # Initialize the current page with any hangover from previous page
            if remainingindex == 0:
                page = remaining[0]
                remainingindex += 1
            
            # Calculate remaining capacity on current page
            remaining_labels_on_page = max_labels_per_page - len(page)
            # Calculate how many more data items can fit (accounting for copies)
            remaining_data_items_on_page = math.ceil(remaining_labels_on_page / copiesperlabel)
            
            # Add labels to current page up to capacity
            for item in remaining[remainingindex : remaining_data_items_on_page]:
                remainingindex += 1
                for i in range(copiesperlabel):
                    page.append(item)
            
            # Identify labels that don't fit on current page
            hangover = page[remaining_labels_on_page:]
            # Trim current page to maximum capacity
            page = page[:remaining_labels_on_page]
            # Add completed page to pages list
            pages.append(page)
            
            # Reset index and update remaining items for next iteration
            remainingindex = 0
            # Combine hangover with remaining unprocessed items
            remaining = hangover + remaining[remainingindex:]
            remainingindex = 0
            
    return firstpage, pages


def format_labels_single(datalist, templatepath, rowindices, columnindices, spec):
    """
    Fills a label template with one label per entry.

    Args:
        datalist (list): Data entries to be printed.
        labelsheetloc (str): Path to the Word label template.
        rowindices (list): List of row indices to use.
        columnindices (list): List of column indices to use.
        copiesperlabel (int): Number of label copies per entry.

    Returns:
        Document: Word document with formatted labels.
    """
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]
    copiesperlabel = spec.copiesperlabel
    textboxformatinput = spec.textboxformatinput
    fontname = spec.fontname
    fontsize = spec.fontsize
    labeldata = 0
    print(labelsheet)
    print("fls")

    for rind in rowindices:
        if labeldata >= len(datalist):
            return labelsheet
 
        
  
        for cind in columnindices:
            if labeldata >= len(datalist):
                print("returned")
                return labelsheet
            print(datalist[labeldata])
            format_label_cell(table.rows[rind].cells[cind], datalist[labeldata], textboxformatinput, fontname, fontsize)
            labeldata += 1
    print("done")
    return labelsheet



def format_labels_firstpage_fromfile(data_list, templatepath, first_page_row_indices, column_indices, first_page_first_row_col_indices, first_page_last_row_col_indices, spec):
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]
    textboxformatinput = spec.textboxformatinput
    fontname = spec.fontname
    fontsize = spec.fontsize

    first_row = first_page_row_indices[0]

    if len(first_page_row_indices) == 1:
        middle_rows = []
        last_row = None
        

    if len(first_page_row_indices) == 2:
        middle_rows = []
        last_row = first_page_row_indices[1]

    if len(first_page_row_indices) > 2:
        middle_rows = first_page_row_indices[1:-1]
        last_row = first_page_row_indices[:-1]

    labelcount = 0
    for cind in first_page_first_row_col_indices:
        if labelcount >= len(data_list):
            return labelsheet
        current_cell = table.rows[first_row].cells[cind]
        format_label_cell(current_cell, data_list[labelcount], textboxformatinput, fontname, fontsize)
        labelcount += 1

    if labelcount >= len(data_list):
        return labelsheet
    if len(middle_rows) > 0:
        for row in middle_rows:
            current_row = table.rows[row]
            for cind in column_indices:
                if labelcount >= len(data_list):
                    return labelsheet
                format_label_cell(current_row.cells[cind], data_list[labelcount], textboxformatinput, fontname, fontsize)
                labelcount += 1
    if first_page_last_row_col_indices != None:
        for cind in first_page_last_row_col_indices:
            if labelcount >= len(data_list):
                return labelsheet
            current_cell = table.rows[last_row].cells[cind]
            format_label_cell(current_cell, data_list[labelcount], textboxformatinput, fontname, fontsize)
            labelcount += 1

    return labelsheet
    


def format_labels_multi(datalist, templatepath, rowindices, columnindices, copiesperlabel, textboxformatinput, fontname, fontsize):
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]
    labelcount = 0
    maxrow_verticalfill = len(rowindices) // copiesperlabel

    for i in range(maxrow_verticalfill):
        rows_to_fill = [table.rows[rowindices[i * copiesperlabel + j]].cells for j in range(copiesperlabel)]
        for cind in columnindices:
            if labelcount >= len(datalist):
                return labelsheet
            for row in rows_to_fill:
                format_label_cell(row[cind], datalist[labelcount], textboxformatinput, fontname, fontsize)
            labelcount += 1

    # Write the last rows of labels
    remainingrows = rowindices[-1] % copiesperlabel

    lastrowcolumnindices = []
    for ind in columnindices:
        if ind % copiesperlabel == 0:
            if (columnindices[-1] - ind) / copiesperlabel >= 1:
                lastrowcolumnindices.append(ind) 

    for rind in rowindices[-remainingrows:]:
        currentrow = table.rows[rind].cells
        for cind in lastrowcolumnindices:
            if labelcount >= len(datalist):
                return labelsheet
            
            cells_to_write = []
            for i in range(copiesperlabel):
                cells_to_write.append(cind + (2 * i))
            for cell in cells_to_write:
                format_label_cell(currentrow[cell], datalist[labelcount], textboxformatinput, fontname, fontsize)
            labelcount += 1
    return labelsheet


def format_labels_identical(text_box_input, templatepath, row_indices, column_indices, fontname, fontsize):
    """
    Formats the entire label sheet for a given text input and applies the label template formatting.

    Args:
        textinput (str): The text to populate each label cell with.
        labelsheetloc (str): Path to the label template (Word document).

    Returns:
        Document: A `docx.Document` object with the populated label data.
    """
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]

    for rind in row_indices:
        currentrow = table.rows[rind].cells

        for cind in column_indices:
            format_label_cell(currentrow[cind], text_box_input, None, fontname, fontsize)


    return labelsheet


def format_labels_incremental(text_box_input, templatepath, row_indices, column_indices, fontname, fontsize):
    """
    Populates and formats the text in the table in the specified label sheet with the sequential numbers starting from 
    the given starting number.

    Args:
        labelsheetpath (str): Path to the Word document that contains the label template.
        starting (str): The starting number.  Supported formats: "prefix-####" (e.g., "25-0001" or "ABC-0001"), "#####" (e.g., "00123"),
        "prefix####" (e.g. "AB0001").

    Returns:
        Document: The modified Word document with the filled labels.
    """
    labelsheet = Document(templatepath)
    table = labelsheet.tables[0]
    # Check if number is numeric or has a prefix or nonalphanumeric character
    delimiter = None
    if text_box_input.isnumeric():
        tubenumber = int(text_box_input)
        cellcontentformat = "{}"
        zfillvar = len(text_box_input)
    else:
        if not text_box_input.isalnum():
            for character in text_box_input:
                if not character.isalnum():
                    if not delimiter:
                        delimiter = character
                    else:
                        sys.exit("Unsupported number format.")
            tubelist = text_box_input.split(delimiter)
            tubenumber = int(tubelist[1])
            prefix = tubelist[0]
            cellcontentformat = prefix + delimiter + "\n{}"
            zfillvar = len(tubelist[1])
        if text_box_input.isalnum():
            tubenumber = ""
            prefix = ""
            prefixended = False
            for character in text_box_input:
                if character.isalpha():
                    if prefixended == False:
                        prefix = prefix + character
                    else:
                        sys.exit("Unsupported number format.")
                if character.isnumeric():
                    prefixended == True
                    tubenumber = tubenumber + character
            tubenumber = int(tubenumber)
            cellcontentformat = prefix + "\n{}"
            zfillvar = len(text_box_input) - len(prefix) + 1

    # Iterate over the rows and columns to fill in the tube numbers
    for rind in row_indices:
        currentrow = table.rows[rind].cells
        for cind in column_indices:
            printnumber = str(tubenumber).zfill(zfillvar)
            format_label_cell(currentrow[cind], cellcontentformat.format(printnumber), None, fontname, fontsize)
            tubenumber +=1

    return labelsheet

def format_label_cell(cell, data, textboxformatinput, fontname, fontsize):
    """
    Populates a single label cell with the given data and formats it according to the label template.

    This function takes the extracted data and fills the corresponding label cell with the entry 
    number, last name, first name, and date.

    Args:
        cell (docx.table._Cell): The label cell in which the data will be inserted.
        data (list): A list containing the entry's label data. The expected structure is:
                 [identifier, last name, first name, date]. The `date` is expected to be a `datetime` object, but the program will also work if it's a str.

    Returns:
        None: The function directly modifies the cell's content and style.

    Example:
        format_label_cell(cell, ["25-0001", "Doe", "John", datetime.datetime(2025, 3, 29)])
        This would fill the cell with the following formatted text:
        "25-0001
         Doe, John
         03/29/2025"
    """
    print("format_label_cell")
    if textboxformatinput:
        label_text = apply_format_to_row(textboxformatinput, data)
        cell.text = label_text

    else: 
        cell.text = data
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(int(fontsize))
            run.font.name = fontname
            run.bold = True
    return

def apply_format_to_row(textboxformatinput, row_data):
    """
    Applies a label format string with placeholders to a row of data.
    Skips missing or None values and formats datetime objects as dates.

    Args:
        textboxformatinput(str): A string with placeholders like "{SampleID}\n{Date}".
        row_data (list): A list of values in the same order as placeholders.

    Returns:
        str: The formatted label string.
    """
    placeholders = re.findall(r'{(.*?)}', textboxformatinput)
    placeholder_to_value = {}

    for i, key in enumerate(placeholders):
        if i >= len(row_data):
            continue
        value = row_data[i]
        if value is None:
            continue
        if isinstance(value, datetime):
            value = value.strftime('%m-%d-%Y')
        placeholder_to_value[key] = str(value)
    try:
        return textboxformatinput.format(**placeholder_to_value)
    except KeyError:
        return ""  # Return empty string if format fails for some reason

def combine_docs(doc1, doc2):
    """
    Appends all content from doc2 into doc1.

    Args:
        doc1 (Document): First document.
        doc2 (Document): Second document to append.

    Returns:
        Document: Combined document.
    """
    print("called combine_docs")
    composer = Composer(doc1)
    composer.append(doc2)
    return composer.doc