"""
Preset file I/O: reading, copying, and deleting the .json preset files that
back the preset dropdown. No Tkinter dependency, so this is testable without
a display.
"""

import json
import os
import shutil


def ensure_builtin_presets_copied(built_in_presets_dir, user_presets_dir):
    """
    Copies every built-in preset .json into the user's preset folder, unless
    a file of the same name is already there.
    """
    for filename in os.listdir(built_in_presets_dir):
        if filename.endswith(".json"):
            src = os.path.join(built_in_presets_dir, filename)
            dest = os.path.join(user_presets_dir, filename)
            if not os.path.exists(dest):
                shutil.copyfile(src, dest)


def load_presets_from_folder(folder):
    """
    Loads every .json preset in folder.

    Returns:
        dict: preset display name -> (file path, preset data dict). If two
        files share a display name, the later one (by directory listing
        order) wins.
    """
    presets = {}
    if os.path.exists(folder):
        for file in os.listdir(folder):
            if file.endswith(".json"):
                path = os.path.join(folder, file)
                with open(path, "r") as f:
                    data = json.load(f)
                    name = data.get("name", file)
                    presets[name] = (path, data)
    return presets


def delete_preset_files(paths):
    """Deletes each preset file in paths."""
    for path in paths:
        os.remove(path)
