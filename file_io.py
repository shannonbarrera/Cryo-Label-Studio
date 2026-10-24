import os
import sys


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
