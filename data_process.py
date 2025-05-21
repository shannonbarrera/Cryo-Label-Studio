import re

def truncate_data(data_list, truncation_indices):
    """
    Truncates specified string fields in a list of lists based on given indices.

    Args:
        data (list of lists): Original data, e.g., [['ID', 'Last', 'First', 'Date'], ...]
        truncation_indices (list of lists): Instructions in format [row_index, start, end]

    Returns:
        list of lists: Modified data with truncated strings
    """
    # Make a deep copy so we don't modify the original list
    new_data = [row[:] for row in data_list]
    for i in range(len(data_list)):
        for instruction in truncation_indices:
            cell_index, start_idx, end_idx = instruction
            target_str = new_data[i][0]
        new_data[i][cell_index]= target_str[start_idx:end_idx]
    return new_data


def is_valid_serial_format(text):
    if len(text) > 12:
        return False
    allowed_formats = [
        r'^\d{1,12}$',                          # Just numbers
        r'^[A-Za-z0-9]{1,5}-\d{1,10}$',         # Prefix + dash + up to 10 digits
        r'^[A-Za-z0-9]{1,5}_\d{1,10}$',         # Prefix + underscore + up to 10 digits
        r'^[A-Za-z0-9]{1,5}\d{1,10}$',          # Prefix directly followed by up to 10 digits
    ]
    return any(re.fullmatch(fmt, text) for fmt in allowed_formats)