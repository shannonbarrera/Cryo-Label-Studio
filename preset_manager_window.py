"""
The "Edit Presets" window: lets the user pick existing presets to edit
(delegating to PresetEditor) or delete.
"""

import tkinter as tk
from tkinter import messagebox

from file_io import resource_path
from preset_store import delete_preset_files


class PresetManagerWindow(tk.Toplevel):
    def __init__(self, master, presets, on_edit, on_presets_changed):
        """
        Args:
            master: parent Tk window.
            presets (dict): name -> (path, data), the current preset list.
            on_edit (callable): called with (name, path, data) when the user
                clicks "Edit" on a selected preset.
            on_presets_changed (callable): called with no args after presets
                are deleted; must return the refreshed presets dict.
        """
        super().__init__(master)
        self.presets = presets
        self.on_edit = on_edit
        self.on_presets_changed = on_presets_changed

        self.title("Edit Presets")
        self.geometry("400x300")
        self.iconbitmap(resource_path("app_icon.ico"))

        self.listbox = tk.Listbox(self, selectmode=tk.MULTIPLE)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._populate_listbox()

        self.status_label = tk.Label(self, text="", fg="green")
        self.status_label.pack(pady=5)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="Edit", command=self._edit_selected).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Delete", command=self._delete_selected).pack(side=tk.LEFT, padx=10)

    def _populate_listbox(self):
        self.listbox.delete(0, tk.END)
        for name in self.presets:
            self.listbox.insert(tk.END, name)

    def _edit_selected(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        name = self.listbox.get(selected[0])
        path, data = self.presets[name]
        self.on_edit(name, path, data)
        self.destroy()

    def _delete_selected(self):
        selected = self.listbox.curselection()
        if not selected:
            return

        confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(selected)} presets?")
        if not confirm:
            return

        deleted_names = [self.listbox.get(i) for i in selected]
        delete_preset_files([self.presets[name][0] for name in deleted_names])

        self.presets = self.on_presets_changed()
        self._populate_listbox()
        self.status_label.config(text=f"Deleted: {', '.join(deleted_names)}")

        self.lift()
        self.focus_force()
