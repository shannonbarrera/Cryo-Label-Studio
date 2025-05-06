
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import csv

class PresetEditor(tk.Toplevel):
    def __init__(self, master, preset_type="File", preset_data=None, preset_path=None, on_save=None):
        super().__init__(master)
        title = "New Text Input Preset" if preset_type == "Text" and preset_data is None else                 "New File Input Preset" if preset_data is None else "Edit Preset"
        self.title(title)
        self.geometry("550x600")

        self.preset_type = preset_type
        self.preset_data = preset_data or {}
        self.preset_path = preset_path
        self.on_save = on_save
        self.entries = {}

        self.header_buttons_frame = None
        self.textbox_format = None

        fields = [
            ("name", "Preset Name"),
            ("labeltemplatepath", "Label Sheet Template"),
            ("inputtype", "Input Type"),
            ("copiesperlabel", "Labels per Sample"),
            ("fontname", "Font Name"),
            ("fontsize", "Font Size"),
            ("outputformat", "Output Format"),
            ("outputfilenameprefix", "Default Output Filename"),
            ("color_theme", "Color Scheme")
        ]

        if self.preset_type == "File":
            fields.append(("partialsheet", "Partial Sheet Selection"))
        elif self.preset_type == "Text":
            fields.insert(2, ("logic", "Logic"))  # Insert after labeltemplatepath
            # Remove "inputtype" and "copiesperlabel" from Text
            fields = [f for f in fields if f[0] not in ("inputtype", "copiesperlabel")]

        row_counter = 0
        for key, label in fields:
            tk.Label(self, text=label).grid(row=row_counter, column=0, sticky="w", padx=10, pady=4)

            if key == "partialsheet":
                var = tk.BooleanVar()
                var.set(self.preset_data.get(key, False))
                cb = tk.Checkbutton(self, variable=var)
                cb.grid(row=row_counter, column=1, padx=10, pady=4, sticky="w")
                self.entries[key] = var

            elif key == "outputformat":
                cb = ttk.Combobox(self, values=["PDF", "DOCX"], state="readonly")
                cb.set(self.preset_data.get(key, "PDF"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "color_theme":
                cb = ttk.Combobox(self, values=["Pink", "Green", "Blue", "Yellow", "Purple"], state="readonly")
                cb.set(self.preset_data.get(key, "Pink"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "labeltemplatepath":
                cb = ttk.Combobox(self, values=["Cryo Dots", "Avery 5160", "ToughSpots 1.5ml"], state="readonly")
                cb.set(self.preset_data.get(key, "Cryo Dots"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "logic":
                cb = ttk.Combobox(self, values=["Identical", "Serial"], state="readonly")
                cb.set(self.preset_data.get(key, "Identical"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "inputtype":
                cb = ttk.Combobox(self, values=["CSV", "XLSX"], state="readonly")
                cb.set(self.preset_data.get(key, "CSV"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "copiesperlabel":
                cb = ttk.Combobox(self, values=[str(i) for i in range(1, 11)], state="readonly")
                cb.set(str(self.preset_data.get(key, "1")))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "fontname":
                cb = ttk.Combobox(self, values=["Arial", "Courier", "Helvetica", "Times", "Verdana"], state="readonly")
                cb.set(self.preset_data.get(key, "Arial"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "fontsize":
                cb = ttk.Combobox(self, values=[str(i) for i in range(6, 21)], state="readonly")
                cb.set(str(self.preset_data.get(key, "10")))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb

            else:
                entry = tk.Entry(self, width=40)
                entry.insert(0, str(self.preset_data.get(key, "")))
                entry.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = entry

            row_counter += 1

        # File input preset: sample file loader and column mapping
        if self.preset_type == "File":
            tk.Button(self, text="Upload Sample File", command=self.load_sample_file).grid(row=row_counter, column=0, columnspan=2, pady=10)
            row_counter += 1

            self.header_buttons_frame = tk.Frame(self)
            self.header_buttons_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="w")
            row_counter += 1

            tk.Label(self, text="Label Format").grid(row=row_counter, column=0, sticky="nw", padx=10)
            self.textbox_format = tk.Text(self, width=40, height=4)
            self.textbox_format.grid(row=row_counter, column=1, padx=10)
            row_counter += 1

        self.save_button = tk.Button(self, text="Save Preset", command=self.save_preset)
        self.save_button.grid(row=row_counter, column=0, columnspan=2, pady=20)

    def load_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            with open(path, newline="") as f:
                reader = csv.DictReader(f)
                headers = reader.fieldnames
                for widget in self.header_buttons_frame.winfo_children():
                    widget.destroy()
                for i, header in enumerate(headers):
                    btn = tk.Button(self.header_buttons_frame, text=header, command=lambda h=header: self.insert_field(h))
                    btn.grid(row=i//3, column=i%3, padx=5, pady=5)

    def insert_field(self, field_name):
        self.textbox_format.insert(tk.INSERT, f"{{{field_name}}}")

    def save_preset(self):
        preset = {"presettype": self.preset_type}
        for key, widget in self.entries.items():
            if isinstance(widget, tk.BooleanVar):
                preset[key] = widget.get()
            elif isinstance(widget, ttk.Combobox):
                preset[key] = widget.get()
            else:
                val = widget.get()
                preset[key] = int(val) if val.isdigit() else val

        if self.textbox_format:
            preset["textboxformatinput"] = self.textbox_format.get("1.0", tk.END).strip()

        if not self.preset_path:
            self.preset_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialdir="presets",
                title="Save Preset As"
            )

        if self.preset_path:
            os.makedirs(os.path.dirname(self.preset_path), exist_ok=True)
            with open(self.preset_path, "w") as f:
                json.dump(preset, f, indent=4)
            messagebox.showinfo("Success", "Preset saved successfully.")
            if self.on_save:
                self.on_save()
            self.destroy()
