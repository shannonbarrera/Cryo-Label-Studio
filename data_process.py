import re


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

