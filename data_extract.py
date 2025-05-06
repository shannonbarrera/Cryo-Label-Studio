import openpyxl as xlsx
import csv


def get_data_list_csv(input_file_path, tablecoords, textboxformatinput):
    """
    Extracts label data from a CSV file using defined table coordinates and format.

    Args:
        input_file_path (str): Path to the CSV file.
        tablecoords (list): List containing two sublists with start and end [row, col] coordinates.
        textboxformatinput (str): Format string describing column layout using letter indices.

    Returns:
        list: Extracted label data.
    """
    label_data_list_format = get_label_data_list_format(textboxformatinput)

    start_row, start_col = tablecoords[0]
    end_row, end_col = tablecoords[1]
    # Convert label_data_list_format_alpha into a list of indices from the row
    batchdata = []
    with open(input_file_path, 'r') as file:
        csv_reader = list(csv.reader(file))
        for row in csv_reader[start_row:end_row + 1]:
            data = []
            for index in label_data_list_format:
                data.append(row[index])
            batchdata.append(data)

    return batchdata


def get_data_list_xlsx(input_file_path, tablecoords, textboxformatinput):
    """
    Extracts label data from an Excel (.xlsx) file.

    Args:
        input_file_path (str): Path to the Excel file.
        tablecoords (list): Table boundaries in the format [[start_row, start_col], [end_row, end_col]].
        textboxformatinput (str): Column layout format using letters.

    Returns:
        list: Extracted label data.
    """
    workbook = xlsx.load_workbook(filename=input_file_path, read_only=True)
    sheet = workbook.active
    info = extract_label_info(sheet, tablecoords, textboxformatinput)
    workbook.close()
    return info

# Read the worksheet and format the label info into a list of lists
def extract_label_info(sheet, tablecoords, textboxformatinput):
    """
    Extracts label data from a worksheet using the defined coordinates and format.

    Args:
        sheet (Worksheet): Active Excel worksheet.
        tablecoords (list): Table boundaries.
        textboxformatinput (str): Format string defining column layout.

    Returns:
        list: List of label entries (each entry is a list).
    """

    start_row, start_col = tablecoords[0]
    end_row, end_col = tablecoords[1]
    label_data_list_format = get_label_data_list_format(textboxformatinput)
    
    batchdata = []
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row,
                               min_col=start_col, max_col=end_col):
        extracted = []
        for idx in label_data_list_format:
            if idx < len(row):
                extracted.append(row[idx].value)
        if extracted:
            batchdata.append(extracted)
    return batchdata

def get_label_data_list_format(textboxformatinput):
    """
    Converts a format string with letter references (e.g., 'B\nD, C\nE') into column indices.

    Args:
        textboxformatinput (str): Format string where letters refer to column indices.

    Returns:
        list: List of zero-based column indices.
    """
    label_data_list_format = [] # Specifies all of the indices to save from each table row 
    for character in textboxformatinput:
        if character.isalpha():
            letter = character.lower()
            number = ord(letter) - ord('a')
            label_data_list_format.append(number)
    return label_data_list_format
