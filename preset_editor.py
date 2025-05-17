
import tkinter as tk
from label_templates import label_templates
from tkinter import filedialog, messagebox, ttk
import json
import os
import csv
import openpyxl as xlsx

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
            ("labeltemplate", "Label Sheet Template"),
            ("copiesperlabel", "Labels per Sample"),
            ("fontname", "Font Name"),
            ("fontsize", "Font Size"),
            ("outputformat", "Output Format"),
            ("outputfilenameprefix", "Default Output Filename"),
            ("output_add_date", "Add Date to Filename"),
            ("color_theme", "Color Scheme")
        ]

        self.template_display_map = {
            v["display_name"]: k for k, v in label_templates.items()
        }
        self.template_internal_map = {
            k: v["display_name"] for k, v in label_templates.items()
        }


        if self.preset_type == "File":
            fields.append(("partialsheet", "Partial Sheet Selection"))
        elif self.preset_type == "Text":
            fields.insert(2, ("logic", "Logic")) 
   
            fields = [f for f in fields if f[0] != "copiesperlabel"]

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

            elif key == "output_add_date":
                var = tk.BooleanVar()
                var.set(self.preset_data.get(key, False))
                cb = tk.Checkbutton(self, variable=var)
                cb.grid(row=row_counter, column=1, padx=10, pady=4, sticky="w")
                self.entries[key] = var


            elif key == "labeltemplate":
                display_names = list(self.template_display_map.keys())
                cb = ttk.Combobox(self, values=display_names, state="readonly")
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)

                # Set display name from internal value in the preset (fallback to first)
                internal_value = self.preset_data.get(key)
                display_name = self.template_internal_map.get(internal_value, display_names[0])
                cb.set(display_name)

                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb


            elif key == "logic":
                cb = ttk.Combobox(self, values=["Identical", "Serial"], state="readonly")
                cb.set(self.preset_data.get(key, "Identical"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                cb.bind("<<ComboboxSelected>>", self.toggle_labels_per_serial)
                self.entries[key] = cb

                # Setup Labels Per Serial widgets
                self.labels_per_serial_row = row_counter + 1
                self.labels_per_serial_label = tk.Label(self, text="Labels Per Serial")
                self.labels_per_serial_dropdown = ttk.Combobox(self, values=[str(i) for i in range(1, 11)], state="readonly")
                self.labels_per_serial_dropdown.set(str(self.preset_data.get("labels_perserial", "1")))

                if cb.get() == "Serial":
                    self.labels_per_serial_label.grid(row=self.labels_per_serial_row, column=0, sticky="w", padx=10, pady=4)
                    self.labels_per_serial_dropdown.grid(row=self.labels_per_serial_row, column=1, padx=10, pady=4)

                row_counter += 2  # use 2 rows if Serial is selected, otherwise will be handled dynamically


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
                cb = ttk.Combobox(self, values=["5.5", "6", "6.5", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"], state="readonly")
                cb.set(str(self.preset_data.get(key, "6")))
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
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
            self.textbox_format = tk.Text(self, width=75, height=8)
            self.textbox_format.grid(row=row_counter, column=1, padx=10)
            if self.preset_data.get("textboxformatinput"):
                self.textbox_format.insert("1.0", self.preset_data["textboxformatinput"])
            row_counter += 1

        self.save_button = tk.Button(self, text="Save Preset", command=self.save_preset)
        self.save_button.grid(row=row_counter, column=0, columnspan=2, pady=20)
        self.update_textbox_size()

    def load_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if path:
            self.lift()
            self.focus_force()

            # Get headers based on file extension
            if path.lower().endswith(".csv"):
                with open(path, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    headers = next(reader)  # first row
            elif path.lower().endswith(".xlsx"):
                wb = xlsx.load_workbook(path, read_only=True)
                sheet = wb.active
                headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            else:
                print("Unsupported file type")
                return

            # Clear previous buttons
            for widget in self.header_buttons_frame.winfo_children():
                widget.destroy()

            # Create new buttons for each header
            filtered_headers = [h for h in headers if h and str(h).strip()]

            for i, header in enumerate(filtered_headers):
                btn = tk.Button(
                    self.header_buttons_frame,
                    text=header,
                    command=lambda h=header: self.insert_field(h)
                )
                btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)



    def insert_field(self, field_name):
        self.textbox_format.insert(tk.INSERT, f"{{{field_name}}}")


    def toggle_labels_per_serial(self, event=None):
        if self.entries["logic"].get() == "Serial":
            self.labels_per_serial_label.grid(row=self.labels_per_serial_row, column=0, sticky="w", padx=10, pady=4)
            self.labels_per_serial_dropdown.grid(row=self.labels_per_serial_row, column=1, padx=10, pady=4)
        else:
            self.labels_per_serial_label.grid_remove()
            self.labels_per_serial_dropdown.grid_remove()

    def update_textbox_size(self, event=None):
        if not self.textbox_format:
            return

        display_name = self.entries["labeltemplate"].get()
        internal_name = self.template_display_map.get(display_name)

        from label_templates import label_templates
        template = label_templates.get(internal_name)
        if not template:
            return

        try:
            font_size = float(self.entries["fontsize"].get())
        except (ValueError, TypeError):
            font_size = template.get("default_font_size", 10)

        chars = template.get("chars_per_line", 45)
        lines = template.get("lines_per_label", 6)
        print(chars)
        print(lines)
        self.textbox_format.config(
            width=max(45, chars),
            height=max(6, lines)
        )


    def save_preset(self):
        preset = {"presettype": self.preset_type}
        if self.preset_type == "Text" and self.entries.get("logic", "").get() == "Serial":
            preset["labels_perserial"] = self.labels_per_serial_dropdown.get()

        for key, widget in self.entries.items():
            if key == "labeltemplate":
                display_value = widget.get()
                internal_value = self.template_display_map.get(display_value, display_value)
                preset[key] = internal_value
            elif isinstance(widget, tk.BooleanVar):
                preset[key] = widget.get()
            elif isinstance(widget, ttk.Combobox):
                val = widget.get()
                preset[key] = int(val) if key == "copiesperlabel" else val
            else:
                val = widget.get()
                preset[key] = int(val) if val.isdigit() else val

        # Add UI layout to the preset if you want static layout elements saved
        if self.preset_type == "Text":
            preset["ui_layout"] = {
                "elements": [
                    {"type": "textbox", "id": "user_input"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        elif self.preset_type == "File":
            preset["ui_layout"] = {
                "elements": [
                    {"type": "button", "id": "upload_file", "label": "Load File"},
                    {"type": "textpreview", "id": "preview_area"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }


        if self.textbox_format:
            preset["textboxformatinput"] = self.textbox_format.get("1.0", "end-1c")
            print(preset["textboxformatinput"])

        if not self.preset_path:
            safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in preset.get("name", "preset")).strip().replace(" ", "_")
            counter = 1
            filename = f"{safe_name}.json"
            full_path = os.path.join("presets", filename)
            while os.path.exists(full_path):
                filename = f"{safe_name}_{counter}.json"
                full_path = os.path.join("presets", filename)
                counter += 1
            self.preset_path = full_path


        if self.preset_path:
            os.makedirs(os.path.dirname(self.preset_path), exist_ok=True)
            with open(self.preset_path, "w") as f:
                json.dump(preset, f, indent=4)
            messagebox.showinfo("Success", "Preset saved successfully.")
            if self.on_save:
                self.on_save()
            self.destroy()