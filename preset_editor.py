
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
        self.preset_type = self._resolve_type(preset_type, preset_data)
        self.preset_data = preset_data or {}
        self.preset_path = preset_path
        self.on_save = on_save
        self.entries = {}
        self.header_buttons_frame = None
        self.textbox_format = None

        self._setup_window()
        self._init_template_maps()
        self._define_fields()
        self._create_fields_ui()
        self._handle_format_input_section()
        self._create_save_button()
        self.update_textbox_size()

    def _resolve_type(self, preset_type, preset_data):
        return preset_data.get("presettype") if preset_data and "presettype" in preset_data else preset_type

    def _setup_window(self):
        title = ("Edit " if self.preset_data else "New ") + ("Text Input Preset" if self.preset_type == "Text" else "File Input Preset")
        self.title(title)
        # self.geometry("550x600")

    def _init_template_maps(self):
        self.template_display_map = {v["display_name"]: k for k, v in label_templates.items()}
        self.template_internal_map = {k: v["display_name"] for k, v in label_templates.items()}

    def _define_fields(self):
        self.fields = [
            ("name", "Preset Name"),
            ("labeltemplate", "Label Sheet Template"),
            ("copiesperlabel", "Labels per Sample"),
            ("fontname", "Font Name"),
            ("fontsize", "Font Size"),
            ("text_alignment", "Label Text Alignment"),
            ("outputformat", "Output Format"),
            ("outputfilenameprefix", "Default Output Filename"),
            ("output_add_date", "Add Date to Filename"),
            ("partialsheet", "Partial Sheet Selection"),
            ("color_theme", "Color Scheme")
        ]
        if self.preset_type == "Text":
            self.fields.insert(2, ("identical_or_incremental", "Logic"))

    def _create_fields_ui(self):
        row_counter = 0
        for key, label in self.fields:
            tk.Label(self, text=label).grid(row=row_counter, column=0, sticky="w", padx=10, pady=4)

            if key == "partialsheet" or key == "output_add_date":
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

            elif key == "text_alignment":
                cb = ttk.Combobox(self, values=["Left", "Center", "Right"], state="readonly")
                cb.set(self.preset_data.get(key, "Left"))
                cb.grid(row=row_counter, column=1, padx=10, pady=4)
                self.entries[key] = cb


            elif key == "labeltemplate":
                display_names = list(self.template_display_map.keys())
                cb = ttk.Combobox(self, values=display_names, state="readonly")
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
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

                self.labels_per_serial_row = row_counter + 1
                self.labels_per_serial_label = tk.Label(self, text="Copies Per Label")
                self.labels_per_serial_dropdown = ttk.Combobox(self, values=[str(i) for i in range(1, 11)], state="readonly")
                self.labels_per_serial_dropdown.set(str(self.preset_data.get("copiesperlabel", "1")))

                if cb.get() == "Incremental":
                    self.labels_per_serial_label.grid(row=self.labels_per_serial_row, column=0, sticky="w", padx=10, pady=4)
                    self.labels_per_serial_dropdown.grid(row=self.labels_per_serial_row, column=1, padx=10, pady=4)

                row_counter += 2
                continue

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

    def _handle_format_input_section(self):
        row_counter = max(widget.grid_info()['row'] for widget in self.winfo_children()) + 1

        if self.preset_type == "File":
            tk.Button(self, text="Upload Sample File", command=self.load_sample_file).grid(row=row_counter, column=0, columnspan=2, pady=10)
            row_counter += 1

            self.header_buttons_frame = tk.Frame(self)
            self.header_buttons_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="w")
            row_counter += 1

            tk.Label(self, text="Label Format:").grid(row=row_counter, column=0, columnspan=2, sticky="n", padx=10, pady=(10, 0))
            row_counter += 1

            self.textbox_format = tk.Text(self, width=75, height=8)
            alignment = self.entries.get("text_alignment")
            if hasattr(alignment, "get"):
                align_value = alignment.get().lower()
                self.textbox_format.tag_configure("align", justify=align_value)
                self.textbox_format.tag_add("align", "1.0", "end")

            self.textbox_format.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=(0, 10))

            if self.preset_data.get("textboxformatinput"):
                self.textbox_format.insert("1.0", self.preset_data["textboxformatinput"])
            self.entries["textboxformatinput"] = self.textbox_format
            row_counter += 1

        elif self.preset_type == "Text":
            tk.Label(self, text="Label Text Format:").grid(row=row_counter, column=0, sticky="nw", padx=10, pady=(10, 0))

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


    def _create_save_button(self):
        self.save_button = tk.Button(self, text="Save Preset", command=self.save_preset)
        self.save_button.grid(row=999, column=0, columnspan=2, pady=20)  # Replace 999 with calculated row

    def save_preset(self):
        preset = {"presettype": self.preset_type}
        preset.update(self._gather_entry_values())
        preset["ui_layout"] = self._get_ui_layout()

        if self.on_save:
            self.on_save(preset, self.preset_path)
        self.destroy()

    def _gather_entry_values(self):
        values = {}
        for key, widget in self.entries.items():
            if key == "labeltemplate":
                display_value = widget.get()
                internal_value = self.template_display_map.get(display_value, display_value)
                values[key] = internal_value

            elif isinstance(widget, tk.BooleanVar):
                values[key] = widget.get()

            elif isinstance(widget, ttk.Combobox):
                val = widget.get()
                values[key] = int(val) if val.isdigit() else val

            elif isinstance(widget, tk.Text):
                val = widget.get("0.0", "end-1c").rstrip()
                values[key] = val

            else:
                try:
                    val = widget.get()
                    values[key] = int(val) if val.isdigit() else val
                except Exception:
                    print(f"Skipping unknown widget for key: {key}")

        return values

    def _get_ui_layout(self):
        if self.preset_type == "Text":
            return {
                "elements": [
                    {"type": "textbox", "id": "user_input"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        elif self.preset_type == "File":
            return {
                "elements": [
                    {"type": "file_upload", "id": "upload_sample"},
                    {"type": "textbox", "id": "textboxformatinput"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        return {}

    def update_textbox_size(self, event=None):
        labeltemplate_widget = self.entries.get("labeltemplate")
        fontsize_widget = self.entries.get("fontsize")

        if hasattr(labeltemplate_widget, "get") and hasattr(fontsize_widget, "get"):
            display_name = labeltemplate_widget.get()
            internal_key = self.template_display_map.get(display_name, display_name)
            template = label_templates.get(internal_key)

            if template:
                chars_per_line = template.get("chars_per_line", 20)
                lines_per_label = template.get("lines_per_label", 3)

                try:
                    font_size = float(fontsize_widget.get())
                except ValueError:
                    font_size = 6

                # Apply size to the box
                if self.textbox_format:
                    self.textbox_format.config(width=chars_per_line, height=lines_per_label)


    def load_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if path:
            self.lift()
            self.focus_force()

            # Get headers based on file extension
            if path.lower().endswith(".csv"):
                with open(path, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    headers = next(reader)
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

            filtered_headers = [h for h in headers if h and str(h).strip()]

            # Set up an inner frame with a fixed width that will be centered by pack
            grid_frame = tk.Frame(self.header_buttons_frame, width=400)
            grid_frame.pack(pady=10)
            grid_frame.pack_propagate(False)  # Prevent it from shrinking to fit

            # Grid buttons inside that fixed-size frame
            for i, header in enumerate(filtered_headers):
                btn = tk.Button(
                    grid_frame,
                    text=header,
                    command=lambda h=header: self.insert_field(h)
                )
                btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)




    
    def insert_field(self, column_name):
        if self.textbox_format:
            self.textbox_format.insert(tk.INSERT, f"{{{column_name}}}")


