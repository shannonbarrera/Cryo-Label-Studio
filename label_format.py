import sys
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from datetime import datetime
from label_templates import label_templates


def get_layout_from_spec(spec):
    template = label_templates.get(spec.labeltemplate)
    if not template:
        raise ValueError(f"Unknown label template ID: {spec.labeltemplate}")

    rows = template["labels_down"]
    cols = template["labels_across"]
    print(rows)
    print(cols)
    row_indices = []
    i = 0
    while len(row_indices) < rows:
        row_indices.append(i)
        i += 2

    col_indices = []
    j = 0
    while len(col_indices) < cols:
        col_indices.append(j)
        j += 2

    print(col_indices)
    print(row_indices)
    return row_indices, col_indices

def generate_blank_docx_template(labeltemplate_id):
    """
    Creates a blank Word Document with a table matching the specified label template layout.

    Args:
        labeltemplate_id (str): Key from label_templates.py (e.g., "LCRY-1700").

    Returns:
        docx.Document: A Word document with the table structure and label cell dimensions.
    """
    if labeltemplate_id not in label_templates:
        raise ValueError(f"Unknown label template: {labeltemplate_id}")

    spec = label_templates[labeltemplate_id]
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(spec["page_width_in"])
    section.page_height = Inches(spec["page_height_in"])
    section.top_margin = Inches(spec["top_margin_in"])
    section.left_margin = Inches(spec["side_margin_in"])
    section.right_margin = Inches(spec["side_margin_in"])
    section.bottom_margin = Inches(0.5)

    rows = (spec["labels_down"] * 2) - 1
    cols = (spec["labels_across"] * 2) - 1
    table = doc.add_table(rows=rows, cols=cols)

    label_width = spec["label_width_in"]
    label_height = spec["label_height_in"]
    spacer_col_width = spec["horizontal_pitch_in"] - label_width
    spacer_row_height = spec["vertical_pitch_in"] - label_height

    for r in range(rows):
        tr = table.rows[r]._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')

        if r % 2 == 0:
            trHeight.set(qn('w:val'), str(int(label_height * 1440)))
        else:
            trHeight.set(qn('w:val'), str(int(spacer_row_height * 1440)))
        trPr.append(trHeight)

        for c in range(cols):
            cell = table.cell(r, c)
            if c % 2 == 0:
                cell.width = Inches(label_width)
            else:
                cell.width = Inches(spacer_col_width)
            # Remove padding
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            for side in ["top", "start", "bottom", "end"]:
                mar = OxmlElement(f'w:tcMar')
                mar.set(qn(f'w:{side}'), "0")
                tcPr.append(mar)

    return doc



def get_max_labels_per_page(spec, labeltemplate, copiesperlabel):
    """
    Calculates how many label entries can fit per page.

    Args:
        labeltemplate (str): Path to label template file.
        labelsheetlayouttype (str): Layout type of the label sheet.
        copiesperlabel (int): Number of repeated labels per entry.

    Returns:
        int: Maximum number of unique label entries per page.
    """
    
    row_indices, column_indices = get_layout_from_spec(spec)
    total_cells = len(row_indices) * len(column_indices)
    return total_cells // copiesperlabel

def format_labels_single(datalist, labeltemplate, rowindices, columnindices, copiesperlabel, textboxformatinput, fontname, fontsize):
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
    labelsheet = Document(labeltemplate)
    table = labelsheet.tables[0]

    labeldata = 0

    for rind in rowindices:
        if labeldata >= len(datalist):
            return labelsheet

        currentrow = table.rows[rind].cells

        for cind in columnindices:
            if labeldata >= len(datalist):
                return labelsheet

            format_label_cell(currentrow[cind], datalist[labeldata], textboxformatinput, fontname, fontsize)
            labeldata += 1

    return labelsheet


def format_labels_multi(datalist, labeltemplate, rowindices, columnindices, copiesperlabel, textboxformatinput, fontname, fontsize):
    """
    Fills a label template with multiple vertical label copies per entry.

    Args:
        datalist (list): List of data entries.
        labeltemplate (str): Template file location.
        rowindices (list): Row indices of the table.
        columnindices (list): Column indices of the table.
        copiesperlabel (int): Number of label copies per entry.

    Returns:
        Document: Word document with formatted multi-copy labels.
    """
    labelsheet = Document(labeltemplate)
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


def format_labels_identical(text_box_input, labeltemplate, row_indices, column_indices, fontname, fontsize):
    """
    Formats the entire label sheet for a given text input and applies the label template formatting.

    Args:
        textinput (str): The text to populate each label cell with.
        labelsheetloc (str): Path to the label template (Word document).

    Returns:
        Document: A `docx.Document` object with the populated label data.
    """
    labelsheet = Document(labeltemplate)
    table = labelsheet.tables[0]

    for rind in row_indices:
        currentrow = table.rows[rind].cells

        for cind in column_indices:
            format_label_cell(currentrow[cind], text_box_input, None, fontname, fontsize)


    return labelsheet


def format_labels_incremental(text_box_input, labeltemplate, row_indices, column_indices, fontname, fontsize):
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
    labelsheet = Document(labeltemplate)
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
    
    if textboxformatinput:
        label_text = apply_format_to_row(textboxformatinput, data)
        

        cell.text = label_text

    else: 
        cell.text = data
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(fontsize)
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
    for element in doc2.element.body:
        doc1.element.body.append(element)
    return doc1
