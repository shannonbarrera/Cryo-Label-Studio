import re
import math


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

def parse_copiesperlabel_input(raw_input):
    """
    Parse a string representing label copy counts into a list of individual values.

    Supports:
    - Single numbers (e.g., "1")
    - Comma-separated lists (e.g., "1, 2, 3")
    - Ranges using dashes (e.g., "1-3")

    Args:
        raw_input (str): The raw input string to parse.

    Returns:
        list of str: A list of individual copy count strings.
    """
    result = []
    parts = [part.strip() for part in raw_input.split(",")]
    for part in parts:
        if "-" in part:
            match = re.match(r"(\d+)-(\d+)", part)
            if match:
                start, end = map(int, match.groups())
                result.extend([str(i) for i in range(start, end + 1)])
        elif part.isdigit():
            result.append(part)
    return result

def estimate_max_chars(label_width_in_inches, font_size_pt, font_name="Arial"):
    """
    Estimate how many characters can fit on a single line based on label width and font size.

    Args:
        label_width_in_inches (float): Width of the label in inches
        font_size_pt (int or float): Font size in points (pt)
        font_name (str): Font name, default Arial

    Returns:
        int: Approximate max number of characters per line
    """
    # Approximate character width in inches for common fonts at given font size
    # 1 pt = 1/72 inch; so 12 pt font = ~0.166 inch high
    avg_char_width_pt = {
        "Arial": 0.5,
        "Courier": 0.6,  # monospaced
        "Helvetica": 0.5,
        "Times": 0.45,
        "Verdana": 0.53
    }

    # get estimated character width in points
    char_width_pt = avg_char_width_pt.get(font_name, 0.5)
    char_width_inches = char_width_pt * font_size_pt / 72.0
    if char_width_inches == 0:
        return 40  # fallback

    return int(label_width_in_inches / char_width_inches)
