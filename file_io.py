import os
import sys
from datetime import datetime

def get_template(labeltemplate):
    # Set the base path for accessing resources
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    # Path to the label template
    labeltemplate_path = os.path.join(base_path, labeltemplate)
    return labeltemplate_path

def get_file_path(output_file_path, outputfilenameprefix, outputformat):
    """
    Determines the file path for saving the generated label document based on the input Excel sheet location.
    The file is named "dnalabels" + date + ".docx".
    Args:
        csvsheetpath (str): Path to the input CSV sheet.
        datalist (list): A list of lists containing extracted data.

    Returns:
        str: The file path for the output label document.
    """
    # Get the current date
    current_date = datetime.now()

    # Format the date without slashes
    formatted_date = current_date.strftime("%y%m%d")

    filepath = os.path.dirname(output_file_path)
    return os.path.join(filepath, f"{outputfilenameprefix}{formatted_date}{outputformat}")


def save_file(filepath, content):
    """
    Saves content to a file, appending a counter to the filename if it already exists.

    Args:
        filepath (str): The desired path to save the file.
        content (str): The content to write to the file.
    """
    filename, extension = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(filepath):
        filepath = f"{filename}_{counter}{extension}"
        counter += 1
    content.save(filepath)
    os.startfile(filepath)

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)

def get_user_presets_folder():
    """
    Returns a user-writable folder for saved presets.
    Creates it if needed.
    """
    base = os.getenv('APPDATA')
    if not base:
        base = os.path.expanduser("~/.CryoLabelStudio")
    folder = os.path.join(base, "CryoLabelStudio", "presets")
    os.makedirs(folder, exist_ok=True)
    return folder
