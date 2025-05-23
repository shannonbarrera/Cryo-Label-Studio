
import tkinter as tk
from label_templates import label_templates
from tkinter import filedialog, messagebox, ttk
import json
import os
import csv
import openpyxl as xlsx

class PresetEditor(tk.Toplevel):
    def __init__(self, master, preset_type="File", preset_data=None, preset_path=None, on_save=None):
        if preset_data and "presettype" in preset_data:
            preset_type = preset_data["presettype"]

        self.preset_type = preset_type
        self.preset_data = preset_data or {}
        self.preset_path = preset_path
        self.on_save = on_save
        self.entries = {}

        if preset_data is None:
            title = "New Text Input Preset" if preset_type == "Text" else "New File Input Preset"
        else:
            title = "Edit Text Input Preset" if preset_type == "Text" else "Edit File Input Preset"

        super().__init__(master)
        self.title(title)
        self.geometry("550x600")

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
            ("partialsheet", "Partial Sheet Selection"),
            ("color_theme", "Color Scheme")
        ]

        self.template_display_map = {
            v["display_name"]: k for k, v in label_templates.items()
        }
        self.template_internal_map = {
            k: v["display_name"] for k, v in label_templates.items()
        }


        if self.preset_type == "Text":
            fields.insert(2, ("identical_or_incremental", "Logic"))

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


            elif key == "identical_or_incremental":
                cb = ttk.Combobox(self, values=["Identical", "Incremental"], state="readonly")
                cb.set(self.preset_data.get(key, "Identical"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                cb.bind("<<ComboboxSelected>>", self.toggle_labels_per_serial)
                self.entries[key] = cb

                # Setup Labels Per Serial widgets
                self.labels_per_serial_row = row_counter + 1
                # NEW — use the same key as File presets
                self.labels_per_serial_label = tk.Label(self, text="Copies Per Label")
                self.labels_per_serial_dropdown = ttk.Combobox(self, values=[str(i) for i in range(1, 11)], state="readonly")
                self.labels_per_serial_dropdown.set(str(self.preset_data.get("copiesperlabel", "1")))


                if cb.get() == "Incremental":
                    self.labels_per_serial_label.grid(row=self.labels_per_serial_row, column=0, sticky="w", padx=10, pady=4)
                    self.labels_per_serial_dropdown.grid(row=self.labels_per_serial_row, column=1, padx=10, pady=4)

                row_counter += 2  # use 2 rows if Serial is selected, otherwise will be handled dynamically


            elif key == "copiesperlabel":
                self.copiesperlabel_row = row_counter
                self.copiesperlabel_label = tk.Label(self, text="Copies Per Label")
                self.copiesperlabel_dropdown = ttk.Combobox(self, values=[str(i) for i in range(1, 11)], state="readonly")
                self.copiesperlabel_dropdown.set(str(self.preset_data.get("copiesperlabel", "1")))

                if self.preset_type == "File" or self.preset_data.get("identical_or_incremental", "Identical") == "Incremental":
                    self.copiesperlabel_label.grid(row=row_counter, column=0, sticky="w", padx=10, pady=4)
                    self.copiesperlabel_dropdown.grid(row=row_counter, column=1, padx=10, pady=4)

                self.entries["copiesperlabel"] = self.copiesperlabel_dropdown


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
            self.entries["textboxformatinput"] = self.textbox_format

            row_counter += 1

        if self.preset_type == "Text":
            # Label Format Section
            format_label = tk.Label(self, text="Label Text Format:")
            format_label.grid(row=row_counter, column=0, sticky="nw", padx=10, pady=(10, 0))

            self.format_entry = tk.Text(self, height=4, width=50)
            self.format_entry.grid(row=row_counter, column=1, padx=10, pady=(10, 0))
            self.entries["textboxformatinput"] = self.format_entry
            self.textbox_format = self.format_entry

            if self.preset_data.get("textboxformatinput"):
                self.textbox_format.insert("1.0", self.preset_data["textboxformatinput"])


            row_counter += 1

            def insert_label_text():
                self.format_entry.insert(tk.INSERT, "{LABEL_TEXT}")

            insert_button = tk.Button(self, text="{LABEL_TEXT}", command=insert_label_text)
            insert_button.grid(row=row_counter, column=1, sticky="w", padx=10, pady=(0, 10))

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
        if self.entries["identical_or_incremental"].get() == "Incremental":
            self.copiesperlabel_label.grid(row=self.copiesperlabel_row, column=0, sticky="w", padx=10, pady=4)
            self.copiesperlabel_dropdown.grid(row=self.copiesperlabel_row, column=1, padx=10, pady=4)
        else:
            self.copiesperlabel_label.grid_remove()
            self.copiesperlabel_dropdown.grid_remove()

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
        self.textbox_format.config(
            width=max(45, chars),
            height=max(6, lines)
        )


    def save_preset(self):
        preset = {"presettype": self.preset_type}

        for key, widget in self.entries.items():
            if key == "labeltemplate":
                # Convert display name back to internal template key
                display_value = widget.get()
                internal_value = self.template_display_map.get(display_value, display_value)
                preset[key] = internal_value

            elif isinstance(widget, tk.BooleanVar):
                preset[key] = widget.get()

            elif isinstance(widget, ttk.Combobox):
                val = widget.get()
                preset[key] = int(val) if val.isdigit() else val

            elif isinstance(widget, tk.Text):
                val = widget.get("0.0", "end-1c").rstrip()
                preset[key] = val

            else:  # Entry or other widget
                try:
                    val = widget.get()
                    preset[key] = int(val) if val.isdigit() else val
                except Exception:
                    print(f"Skipping unknown widget for key: {key}")

        # Always save UI layout (optional, customizable per preset type)
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
                    {"type": "textpreview", "id": "preview_area"},
                    {"type": "button", "id": "upload_file", "label": "Load File"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }

        # Generate filename if not provided
        if not self.preset_path:
            safe_name = "".join(
                c if c.isalnum() or c in (" ", "-", "_") else "_" for c in preset.get("name", "preset")
            ).strip().replace(" ", "_")

            counter = 1
            filename = f"{safe_name}.json"
            full_path = os.path.join("presets", filename)

            while os.path.exists(full_path):
                filename = f"{safe_name}_{counter}.json"
                full_path = os.path.join("presets", filename)
                counter += 1

            self.preset_path = full_path

        # Write to file
        if self.preset_path:
            os.makedirs(os.path.dirname(self.preset_path), exist_ok=True)
            with open(self.preset_path, "w") as f:
                json.dump(preset, f, indent=4)

            messagebox.showinfo("Success", "Preset saved successfully.")
            if self.on_save:
                self.on_save()
            self.destroy()
